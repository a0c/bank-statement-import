# -*- coding: utf-8 -*-
import chardet
import csv
import json
from datetime import datetime
import logging
import re
from StringIO import StringIO

from openerp import api, fields, models
from openerp.exceptions import ValidationError, Warning
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT
from openerp.tools.safe_eval import safe_eval
from openerp.tools.translate import _

_logger = logging.getLogger(__name__)


def is_csv(filename):
    return filename and (filename.lower().endswith('.csv') or filename.lower().endswith('.txt'))


def read_csv(data_file, csv_format=None):
    fieldnames = csv_format and csv_format.header and csv_format.header.split('\n') or None
    delimiter = csv_format and str(csv_format.delimiter) or ';'
    res = csv.DictReader(StringIO(data_file), fieldnames=fieldnames, delimiter=delimiter)
    if len(res.fieldnames) == 1:  # try comma delimiter (COOP)
        res = csv.DictReader(StringIO(data_file), fieldnames=fieldnames, delimiter=',')
    return res


def get_csv_header_and_first_line(data_file, csv_format):
    csv = read_csv(data_file, csv_format)
    csv_header = to_unicode(csv.fieldnames)
    try:
        line = csv.next()
    except StopIteration:
        line = dict(zip(csv_header, csv_header))  # when only 1 line in CSV, it is already consumed by csv.fieldnames
    return csv_header, line


def to_unicode(fieldnames):
    if isinstance(fieldnames, basestring):
        return fieldnames.decode('utf-8')
    return [f.decode('utf-8') for f in fieldnames]


def encode_unicode(fieldnames):
    if isinstance(fieldnames, basestring):
        return fieldnames.encode('utf-8')
    return [f and f.encode('utf-8') or f for f in fieldnames]


RE_LITERAL = re.compile("""^["']([\w%]+)["']$""")


def value_of(column, tx):
    """ Example: "DE83" + Bank label + Account number """
    res = []
    column = encode_unicode(column or '')  # column can be missing in CSV Format
    for col in re.split('\s*\+\s*', column):
        if not col: continue
        match = RE_LITERAL.match(col)
        if match:
            res.append(to_unicode(match.group(1).strip()))
        else:
            res.append(to_unicode(tx.get(col, '')))  # column can be missing in CSV
    return u''.join(res)


def actual_columns(column):
    return [col for col in re.split('\s*\+\s*', column or '') if col and not RE_LITERAL.match(col)]


class AccountBankStatementImport(models.TransientModel):
    _inherit = 'account.bank.statement.import'

    note = fields.Html(readonly=1)
    csv_formats = fields.Char(readonly=1)

    @api.multi
    def import_file(self):
        try:
            return super(AccountBankStatementImport, self).import_file()
        except Warning, e:
            if len(e.args) == 2:
                title, message = e.args
                if title.startswith('Multiple CSV Formats matched:'):
                    parts = title.split(':', 1)
                    self.csv_formats = parts[1].strip()
                    self.note = '%s:<br>%s' % (parts[0], message)
                    return self.action_reload()
                elif title == 'No CSV Formats matched':
                    self.csv_formats = '[]'
                    self.note = '%s:<br>%s' % (title, message)
                    return self.action_reload()
            raise

    def action_reload(self):
        return self.action_view('account_bank_statement_import.account_bank_statement_import_view',
                                'account_bank_statement_import.action_account_bank_statement_import')

    def decode_data_file(self, data_file):
        """ utf-8-sig:    get rid of BOM in SEB EE
            iso-8859-1:   COOP is in ISO-8859 encoding """
        detected_encoding = chardet.detect(data_file)
        for encoding in filter(None, (detected_encoding['encoding'], "utf-8-sig", "iso-8859-1")):
            try:
                return data_file.decode(encoding).encode("utf-8")
            except UnicodeDecodeError:
                pass
        raise Warning('Unknown Encoding', 'File encoding could not be determined')

    def find_matched_csv_format(self, data_file):
        filename = self.env.context.get('filename')
        if is_csv(filename):
            data_file = self.decode_data_file(data_file)
            all_csv_formats = self.env['account.bank.statement.import.csv.format'].search([])
            csv_formats = all_csv_formats.matches(data_file)  # find all csv_formats that match: e.g. only accept UMSATZ.txt from ZIP file
            if len(csv_formats) == 1:
                return csv_formats, read_csv(data_file, csv_formats)  # re-read CSV with the matched csv_format (header)
            elif len(csv_formats) > 1:
                raise Warning("Multiple CSV Formats matched: %s" % (csv_formats.ids,),
                              "File <b>%s</b> matches multiple CSV Formats. Please update <b>File Match Condition</b> "
                              "in CSV Formats to ensure a single match." % (filename, ))
            else:
                if all_csv_formats.skips(data_file):  # if any format skips the file
                    return [None], [None]  # skip file
                raise Warning("No CSV Formats matched",
                              "File <b>%s</b> doesn't match any CSV Format. Please update <b>File Match Condition</b> "
                              "in CSV Formats to ensure a single match. Or copy-paste <b><code>filename_has('%s')"
                              "</code></b> into <b>File Skip Condition</b> to mute warnings for file <b>%s</b>."
                              % (filename, filename, filename))
        return False, False

    @api.model
    def _parse_file(self, data_file):
        fmt, csv = self.find_matched_csv_format(data_file)
        if not csv:
            return super(AccountBankStatementImport, self)._parse_file(data_file)
        if csv == [None]:  # file skipped using fmt.skip_condition
            return []

        transactions = []
        _logger.info('Importing Bank Statement using "%s" CSV Format' % fmt.name)
        fmt.validate_header_against_required_format_columns(csv.fieldnames)
        account_number = currency_code = balance_start = balance_end = False
        try:
            for transaction in csv:
                def val(col): return value_of(col, transaction)
                # skip header line with human-readable column names (e.g. Alfa-Bank)
                if len(val(fmt.drcr)) > 1:
                    continue

                account_number = fmt.parse_account_number(val, account_number)
                currency_code = fmt.parse_currency_code(val, currency_code)

                amount, tx_type = fmt.parse_amount(val), val(fmt.tx_type)
                balance_start, balance_end, is_loaded = fmt.load_balance(tx_type, amount, balance_start, balance_end)
                if is_loaded:
                    continue

                (payee, payee_acc_nr), memo = fmt.parse_payee(val, amount), fmt.parse_memo(val)

                # If bank doesn't provide account numbers, we'll have
                # to find res.partner and res.partner.bank here
                # (normal behavious is to provide 'account_number', which the
                # generic module uses to find partner/bank)
                bank_account_id = partner_id = False
                if payee:
                    banks = self.env['res.partner.bank'].search([('owner_name', '=', payee)], limit=1)
                    if banks:
                        bank_account_id = banks.id
                        partner_id = banks.partner_id.id
                date, ref, archive_id = fmt.parse_date(val), val(fmt.ref), val(fmt.archive_id)
                vals_line = {
                    'date': date,
                    'name': ': '.join(filter(None, (payee, memo or ref))),
                    'ref': ref,
                    'amount': amount,
                    'unique_import_id': '%s-%s-%s-%s-%s' % (archive_id, payee, memo, date, amount),
                    'bank_account_id': bank_account_id,
                    'partner_id': partner_id,
                    'partner_name': payee,
                }
                # 'account_number' will be used for creating res.partner.bank if not found above (bank_account_id)
                if payee_acc_nr:
                    vals_line['account_number'] = payee_acc_nr
                # Memo, reference and payee are not required fields in
                # CSV (although typically at least memo or reference are
                # requested by banks).
                # But the 'name' field of account.bank.statement.line is
                # required=True, so we must always have a value !
                # Fields TX_TYPE and ARCHIVE_ID are typically present in
                # CSV, although not required.
                if not vals_line['name']:  # should never happen
                    vals_line['name'] = ' '.join(filter(None, (tx_type or archive_id, val(fmt.doc_nr))))
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

    @api.multi
    def action_ambiguous_csv_formats(self):
        csv_formats = self.env['account.bank.statement.import.csv.format'].browse(json.loads(self.csv_formats))
        csv_formats = csv_formats or csv_formats.search([])
        return csv_formats.action_view(None, action_xmlid='account_bank_statement_import_csv.action_bank_statement_import_csv_format')


class AccountBankStatementImportCSVFormat(models.Model):
    _name = 'account.bank.statement.import.csv.format'

    name = fields.Char(required=1)
    header = fields.Text(help="If CSV file doesn't have a Header with column names, you need to set it here manually")
    account_number = fields.Char('Account', required=1)
    currency_code = fields.Char('Currency', required=1)
    date = fields.Char(required=1)
    payee = fields.Char('Beneficiary/Payer', required=1)
    payee_account_number = fields.Char("Beneficiary/Payer's Account")
    payer = fields.Char('Payer', help="If Beneficiary and Payer are split into different columns, type Payer column "
                                      "here and Credit transactions will use it instead of Beneficiary/Payer column")
    payer_account_number = fields.Char("Payer's Account")
    memo = fields.Char('Details/Memo', required=1)
    drcr = fields.Char('Debit/Credit')
    amount = fields.Char(required=1)
    ref = fields.Char('Reference Number')
    archive_id = fields.Char('Transfer Reference', help="Archive ID")
    tx_type = fields.Char('Transaction Type')
    tx_type_balance_start = fields.Char('Transaction Type: Balance Start')
    tx_type_balance_end = fields.Char('Transaction Type: Balance End')
    tx_type_turnover = fields.Char('Transaction Type: Turnover')
    doc_nr = fields.Char('Document Number')

    match_condition = fields.Char('File Match Condition', required=1)
    skip_condition = fields.Char('File Skip Condition')

    date_format = fields.Char('Date Format', required=1)
    delimiter = fields.Selection([(',', 'Comma'), (';', 'Semicolon'), ('\t', 'Tab'), (' ', 'Space')], default=';')

    @api.constrains('header')
    def _validate_header(self):
        if self.header:
            columns = self.header.split('\n')
            if len(columns) != len(set(columns)):
                raise ValidationError("Header columns must have different names: duplicates detected.")

    def get_required_columns(self):
        required = [self.account_number, self.currency_code, self.date, self.payee, self.memo, self.amount]
        required_if_set = filter(None, [self.drcr, self.ref, self.archive_id])
        return encode_unicode([col for req in required + required_if_set for col in actual_columns(req)])

    def validate_header_against_required_format_columns(self, header):
        required_columns = self.get_required_columns()
        missing_columns = ["'%s'" % column for column in required_columns if column not in header]
        if missing_columns:
            raise ValidationError("Column(s) %s are missing from CSV file. Actual columns: %s"
                                  % (', '.join(missing_columns), ', '.join("'%s'" % x for x in header)))

    def filename_has(self, *names):
        filename = self.env.context.get('filename')
        return filename and any(name in filename for name in names)

    def is_hex(self, line):
        try:
            int(line, 16)
            return True
        except ValueError:
            return False

    def matches(self, data_file):
        num_cols_csv = len(read_csv(data_file).fieldnames)  # first, read CSV without csv_format (header) - just to count columns

        def match(x):
            if x.header and len(x.header.split('\n')) != num_cols_csv:
                return False
            csv_header, line = get_csv_header_and_first_line(data_file, x)  # then, re-read CSV with the current csv_format (header)
            # re-reading allows referencing cells by column names in match_condition in case when x.header is set, e.g. "line.get(self.drcr) in ('C', 'D')"
            return bool(safe_eval(x.match_condition, locals_dict={
                'self': x, 'csv_header': csv_header, 'line': line, 'header': x.header,
                'filename_has': x.filename_has, 'filename': x.env.context.get('filename'),
                'is_hex': x.is_hex}))

        return self.filtered(match)

    def skips(self, data_file):

        def skip(x):
            if not x.skip_condition:
                return False
            csv_header, line = get_csv_header_and_first_line(data_file, x)  # read CSV with the current csv_format (header)
            return bool(safe_eval(x.skip_condition, locals_dict={
                'self': x, 'csv_header': csv_header, 'line': line, 'header': x.header,
                'filename_has': x.filename_has, 'filename': x.env.context.get('filename')}))

        return self.filtered(skip)

    def resolve_account_number(self, account_number):
        if '%' in account_number:
            account_number = account_number.replace('%', '_')
            bank_account = self.env['res.partner.bank'].search([('acc_number', 'like', account_number)], limit=1)
            if bank_account:
                account_number = bank_account.sanitized_acc_number
        return account_number

    def parse_account_number(self, val, account_number):
        return account_number or self.resolve_account_number(val(self.account_number).lstrip("'"))

    def parse_currency_code(self, val, currency_code):
        return currency_code or val(self.currency_code)

    def parse_date(self, val):
        return datetime.strptime(val(self.date), self.date_format).strftime(DEFAULT_SERVER_DATE_FORMAT)

    def parse_amount(self, val):
        """ drcr not required: if no drcr column - take amount as is """
        amount, drcr = val(self.amount), val(self.drcr)
        amount = float(amount.replace(',', '.'))
        if not self.drcr:
            return amount
        return abs(amount) * (-1.00 if drcr == 'D' else 1)  # abs(): SWEDBANK sends both +-amounts and drcr

    def parse_payee(self, val, amount):
        if self.payer and amount > 0:  # drcr == 'C'
            col, col_nr = self.payer, self.payer_account_number
        else:
            col, col_nr = self.payee, self.payee_account_number
        return val(col).lstrip("'"), val(col_nr).lstrip("'")

    def parse_memo(self, val):
        return val(self.memo).lstrip("'")

    def load_balance(self, tx_type, amount, balance_start, balance_end):
        is_loaded = False
        if tx_type:
            if self.tx_type_balance_start and tx_type == self.tx_type_balance_start:
                balance_start = amount
                is_loaded = True
            if self.tx_type_balance_end and tx_type == self.tx_type_balance_end:
                balance_end = amount
                is_loaded = True
            if self.tx_type_turnover and tx_type == self.tx_type_turnover:
                is_loaded = True
        return balance_start, balance_end, is_loaded
