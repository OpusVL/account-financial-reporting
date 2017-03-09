# -*- coding: utf-8 -*-
# Author: Damien Crier
# Author: Julien Coux
# Copyright 2016 Camptocamp SA
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl.html).

from operator import attrgetter

from . import abstract_report_xlsx
from openerp.report import report_sxw
from openerp import _

from openerp.addons.account_financial_report_webkit.report.general_ledger \
    import GeneralLedgerWebkit

class LineFunData(object):
    """This provides computed and running data to lambda functions used as fields in _get_report_columns()

    Instantiate a new one of these at the beginning of processing of each account.

    Send new_line(line) to that instance for each line in order to keep track of the running totals
    and provide the correct line to the lambda functions.

    After calling new_line(line), the cumul_balance and cumul_balance_curr attributes are updated.

    Instance variables:
      cumul_balance: The current cumulative balance
      cumul_debit: Running total of debit lines
      cumul_credit: Running total of credit lines
    """
    def __init__(self, account, localcontext):
        self.account = account
        self.localcontext = localcontext
        _p = self.localcontext
        # display_initial_balance - direct lift from general_ledger_xls.py from OCA/account-financial-reporting
        self.display_initial_balance = _p['init_balance'][account.id] and \
                (_p['init_balance'][account.id].get(
                    'debit', 0.0) != 0.0 or
                    _p['init_balance'][account.id].get('credit', 0.0) != 0.0)
        initb = _p['init_balance'][account.id]
        self.initial_balance = initb.get('init_balance') or 0.0
        self.initial_debit = initb.get('debit', 0.0)
        self.initial_credit = initb.get('credit', 0.0)
        self.cumul_debit = self.initial_debit
        self.cumul_credit = self.initial_credit
        self.cumul_balance = initb.get('init_balance') or 0.0
        self.cumul_balance_curr = initb.get('init_balance_currency') or 0.0

    def new_line(self, line):
        """Let us know we have a new line, and 
        """
        self.line = line
        self.cumul_debit += line.get('debit') or 0.0
        self.cumul_credit += line.get('credit') or 0.0
        self.cumul_balance_curr += line.get('amount_currency') or 0.0
        self.cumul_balance += line.get('balance') or 0.0

    def finish(self):
        self.final_debit = self.cumul_debit
        self.final_credit = self.cumul_credit
        self.final_balance = self.cumul_balance


def account_code(fundata):
    """-> fundata.account.code"""
    return fundata.account.code

def line_label(fundata):
    """Return the Label text"""
    line = fundata.line
    label_elements = [line.get('lname') or '']
    if line.get('invoice_number'):
        label_elements.append(
            "(%s)" % (line['invoice_number'],))
    label = ' '.join(label_elements)
    return label


class GeneralLedgerXslx(abstract_report_xlsx.AbstractReportXslx):

    def __init__(self, name, table, rml=False, parser=False, header=True,
                 store=False):
        super(GeneralLedgerXslx, self).__init__(
            name, table, rml, parser, header, store)

    def _get_report_name(self):
        return _('General Ledger')

    def _get_report_columns(self, report):
        data = self.parser_instance.localcontext
        return {
            0: {'header': _('Date'), 'field': 'ldate', 'width': 11},
            1: {'header': _('Period'), 'field': 'period_code', 'width': 11},
            2: {'header': _('Entry'), 'field': 'move_name', 'width': 18},
            3: {'header': _('Journal'), 'field': 'jcode', 'width': 8},
            4: {'header': _('Account'), 'field': account_code, 'width': 9},
            5: {'header': _('Partner'), 'field': 'partner_name', 'width': 25},
            6: {'header': _('Label'), 'field': line_label, 'width': 40},
            7: {'header': _('Counterpart'), 'field': 'counterparts', 'width': 25},
            8: {'header': _('Debit'),
                'field': 'debit',
                'field_initial_balance': attrgetter( 'initial_debit' ),
                'field_final_balance': attrgetter( 'final_debit' ),
                'type': 'amount',
                'width': 14},
            9: {'header': _('Credit'),
                'field': 'credit',
                'field_initial_balance': attrgetter( 'initial_credit' ),
                'field_final_balance': attrgetter( 'final_credit' ),
                'type': 'amount',
                'width': 14},
            10: {'header': _('Cumul. Bal.'),
                 'field': attrgetter('cumul_balance'),
                 'field_initial_balance': attrgetter( 'initial_balance' ),
                 'field_final_balance': attrgetter( 'final_balance' ),
                 'type': 'amount',
                 'width': 14},
            # 12: {'header': _('Cur.'), 'field': 'currency_name', 'width': 7},
            # 13: {'header': _('Amount cur.'),
            #      'field': 'amount_currency',
            #      'type': 'amount',
            #      'width': 14},
        }

    def _get_report_filters(self, report):
        form = self.report_data['form']
        return [
            [_('Date range filter'),
                _('From: %s To: %s') % (form['date_from'], form['date_to'])],
            [_('Target Moves'),
                _('All posted entries') if form['target_move'] == 'posted'
                else _('All entries')],
            # [_('Account balance at 0 filter'),
            #     _('Hide') if form['hide_account_balance_at_0'] else _('Show')],
            [_('Centralize filter'),
                _('Yes') if form['centralize'] else _('No')],
        ]

    def _get_col_count_filter_name(self):
        return 2

    def _get_col_count_filter_value(self):
        return 2

    def _get_col_pos_initial_balance_label(self):
        return 5

    def _get_col_count_final_balance_name(self):
        return 5

    def _get_col_pos_final_balance_label(self):
        return 5

    def _generate_report_content(self, workbook, report):
        # import pprint
        # with open('/tmp/localcontext.py', 'wb') as f:
        #     pprint.pprint(self.parser_instance.localcontext, f)

        # with open('/tmp/report_data.py', 'wb') as f:
        #     pprint.pprint(self.report_data, f)
        # For each account
        #for account in report.account_ids:
        #import pdb ; pdb.set_trace()
        data = self.parser_instance.localcontext
        for account in data['objects']:
            fundata = LineFunData(account, data)

            # Write account title
            self.write_array_title(account.code + ' - ' + account.name)

            # TODO What's this for???
            #if not account.partner_ids:
            # Display array header for move lines
            self.write_array_header()

            # Display initial balance line for account
            if fundata.display_initial_balance:
                self.write_initial_balance(account, _('Initial balance'), fundata)

            # Display account move lines
            #for line in account.move_line_ids:
            for line in data['ledger_lines'][account.id]:
                fundata.new_line(line)
                self.write_line(line, fundata)
            fundata.finish()

            # else:
            #     # For each partner
            #     for partner in account.partner_ids:
            #         # Write partner title
            #         self.write_array_title(partner.name)

            #         # Display array header for move lines
            #         self.write_array_header()

            #         # Display initial balance line for partner
            #         self.write_initial_balance(partner, _('Initial balance'))

            #         # Display account move lines
            #         for line in partner.move_line_ids:
            #             self.write_line(line)

            #         # Display ending balance line for partner
            #         self.write_ending_balance(partner, 'partner')

            #         # Line break
            #         self.row_pos += 1

            # Display ending balance line for account
            self.write_ending_balance(account, 'account', fundata)

            # 2 lines break
            self.row_pos += 2

    def write_ending_balance(self, my_object, type_object, fundata):
        """Specific function to write ending balance for General Ledger"""
        if type_object == 'partner':
            name = my_object.name
            label = _('Partner ending balance')
        elif type_object == 'account':
            name = my_object.code + ' - ' + my_object.name
            label = _('Ending balance')
        super(GeneralLedgerXslx, self).write_ending_balance(
            my_object, name, label, fundata
        )


GeneralLedgerXslx(
    'report.account_financial_report_webkit_xlsx.report_general_ledger_xlsx',
    'account.account',
    parser=GeneralLedgerWebkit,  # The parser provides us with the report data via self.parser_instance.localcontext
)
