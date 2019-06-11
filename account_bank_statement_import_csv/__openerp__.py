# -*- coding: utf-8 -*-
{
    'name': 'Import CSV Bank Statement',
    'category': 'Banking addons',
    'version': '8.0.1.0.1',
    'license': 'AGPL-3',
    'author': 'Anton Chepurov @ AVANSER LLC',
    'summary': 'Import CSV Bank Statement',
    'description': """
Import CSV Bank Statement
=========================
""",
    'depends': [
        'account_bank_statement_import'
    ],
    'external_dependencies': {'python': ['chardet']},
    'data': [
        'views/account_bank_statement_import_view.xml',
        'views/account_bank_statement_import_csv_format.xml',
        'security/ir.model.access.csv',
        'data/csv_format.xml',
    ],
    'qweb': [
        'static/src/xml/account_bank_statement_reconciliation.xml',
    ],
    'auto_install': False,
    'installable': True,
}
