# -*- coding: utf-8 -*-
import csv
from datetime import datetime
import StringIO

from openerp import api, models
from openerp.exceptions import ValidationError, Warning
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT
from openerp.tools.translate import _


def get_account_number(transaction, acc_nr, account_number):
    return account_number or transaction[acc_nr].lstrip("'")


def get_currency_code(transaction, curr_code, currency_code):
    return currency_code or transaction[curr_code]


def parse_date(date):
    return datetime.strptime(date, '%d.%m.%Y').strftime(DEFAULT_SERVER_DATE_FORMAT)


SEB_EE_ACC_NR = 'Saaja/maksja konto'


def get_bank_columns(fields):
    """ account_number, currency_code, date, payee, memo, amount, ref,
        Dr/Cr, archive_id, trn_type, doc_nr, payee_acc_nr """
    columns, bank, account_number, first_field = False, False, False, fields[0]
    # Swedbank EN
    if first_field == 'Client account':
        bank = 'Swedbank EN'
        columns = 'Currency', 'Date', 'Beneficiary/Payer', 'Details', 'Amount', 'Reference number',\
                  'Debit/Credit', 'Transfer reference', 'Transaction type', 'Document number'
    # Swedbank RU
    elif first_field == 'Счëт клиентa':
        bank = 'Swedbank RU'
        columns = 'Currency', 'Дата', 'Получатель/Плательщик', 'Пояснение', 'Сумма', 'Номер ссылки',\
                  'Дебит/Кредит', 'Архивный признак', 'Тип сделки', 'Номер документа'
    # Swedbank EE
    elif first_field == 'Kliendi konto' and SEB_EE_ACC_NR not in fields:
        bank = 'Swedbank EE'
        columns = 'Valuuta', 'Kuupäev', 'Saaja/Maksja', 'Selgitus', 'Summa', 'Viitenumber',\
                  'Deebet/Kreedit', 'Arhiveerimistunnus', 'Tehingu tüüp', 'Dokumendi number'
    # SEB EN
    elif first_field == 'Account':
        bank = 'SEB EN'
        account_number = "Beneficiary's account"
        columns = 'Currency', 'Date', "Beneficiary's name", 'Description', 'Amount', 'Reference no.',\
                  '(D/C)', 'Archive ID', 'Type', 'Document No.'
    # SEB RU
    elif first_field == 'Название счета':
        bank = 'SEB RU'
        account_number = "Счёт получателя"
        columns = 'Валюта', 'Дата', 'Наименование получателя', 'Описание', 'Сумма', 'Номер ссылки', \
                  '(D/C)', 'Признак архивации', 'тип', 'Номер документа'
    # SEB EE
    elif first_field == 'Kliendi konto' and SEB_EE_ACC_NR in fields:
        bank = 'SEB EE'
        account_number = SEB_EE_ACC_NR
        columns = 'Valuuta', 'Kuupäev', 'Saaja/maksja nimi', 'Selgitus', 'Summa', 'Viitenumber', \
                  'Deebet/Kreedit (D/C)', 'Arhiveerimistunnus', 'Tüüp', 'Dokumendi number'
    if not columns:
        raise ValidationError('Cannot recognize columns in CSV file')
    missing_columns = ["'%s'" % column for column in columns if column not in fields]
    if missing_columns:
        raise ValidationError("%s recognized by first column '%s', but column(s) %s are missing from CSV file. "
                              "Actual columns: %s"
                              % (bank, first_field, ', '.join(missing_columns), ', '.join("'%s'" % x for x in fields)))
    return (first_field,) + columns + (account_number,)


class AccountBankStatementImport(models.TransientModel):
    _inherit = 'account.bank.statement.import'

    @api.model
    def _check_csv(self, data_file):
        filename = self.env.context.get('filename')
        if filename and filename.lower().endswith('.csv'):
            return csv.DictReader(StringIO.StringIO(data_file), delimiter=';')
        return False

    @api.model
    def _parse_file(self, data_file):
        data_file = data_file.decode("utf-8-sig").encode("utf-8")  # get rid of BOM in SEB EE
        csv = self._check_csv(data_file)
        if not csv:
            return super(AccountBankStatementImport, self)._parse_file(data_file)

        transactions = []
        acc_nr, curr_code, date, payee, memo, amount, ref, drcr, archive_id, trn_type, doc_nr, payee_acc_nr = get_bank_columns(csv.fieldnames)
        account_number = currency_code = balance_start = balance_end = False
        try:
            for transaction in csv:
                account_number = get_account_number(transaction, acc_nr, account_number)
                currency_code = get_currency_code(transaction, curr_code, currency_code)

                _amount = float(transaction[amount].replace(',', '.')) * (-1.00 if transaction[drcr] == 'D' else 1)
                if transaction[trn_type] == 'AS':
                    balance_start = _amount
                    continue
                if transaction[trn_type] == 'LS':
                    balance_end = _amount
                    continue
                if transaction[trn_type] == 'K2':
                    continue

                payee_tx = transaction[payee].lstrip("'")
                memo_tx = transaction[memo].lstrip("'")

                # Since only SEB provides account numbers, we'll have
                # to find res.partner and res.partner.bank here
                # (normal behavious is to provide 'account_number', which the
                # generic module uses to find partner/bank)
                bank_account_id = partner_id = False
                if payee_tx:
                    banks = self.env['res.partner.bank'].search([('owner_name', '=', payee_tx)], limit=1)
                    if banks:
                        bank_account_id = banks.id
                        partner_id = banks.partner_id.id
                vals_line = {
                    'date': parse_date(transaction[date]),
                    'name': ': '.join(filter(None, (payee_tx, memo_tx or transaction[ref]))),
                    'ref': transaction[ref],
                    'amount': _amount,
                    'unique_import_id': '%s-%s-%s' % (transaction[archive_id], payee_tx, memo_tx),
                    'bank_account_id': bank_account_id,
                    'partner_id': partner_id,
                    'partner_name': payee_tx,
                }
                if payee_acc_nr:
                    vals_line['account_number'] = transaction[payee_acc_nr]
                # Memo, reference and payee are not required fields in
                # CSV (although typically at least memo or reference are
                # requested by banks).
                # But the 'name' field of account.bank.statement.line is
                # required=True, so we must always have a value !
                # The field TRNTYPE is a required field in CSV.
                if not vals_line['name']:  # should never happen
                    vals_line['name'] = ' '.join(filter(None, (transaction[trn_type], transaction[doc_nr])))
                transactions.append(vals_line)
        except Exception, e:
            raise Warning(_(
                "The following problem occurred during import. "
                "The file might not be valid.\n\n %s" % e.message
            ))

        vals_bank_statement = {
            'transactions': transactions,
            'balance_start': balance_start,
            'balance_end_real': balance_end,
        }
        return currency_code, account_number, [vals_bank_statement]

    @api.model
    def _find_currency_id(self, currency_code):
        """ override to search for company-specific currency first, then for no-company currency,
            and only then for any currency that matches the currency_code """
        if currency_code:
            currency_ids = \
                self.env['res.currency'].search([('name', '=ilike', currency_code), ('company_id', '=', self.env.user.company_id.id)]) or \
                self.env['res.currency'].search([('name', '=ilike', currency_code), ('company_id', '=', False)]) or \
                self.env['res.currency'].search([('name', '=ilike', currency_code)])
            if currency_ids:
                return currency_ids[0].id
            else:
                raise Warning(_(
                    'Statement has invalid currency code %s') % currency_code)
        # if no currency_code is provided, we'll use the company currency
        return self.env.user.company_id.currency_id.id
