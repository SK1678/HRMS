# -*- coding: utf-8 -*-
"""
import_log.py
--------------
Persistent model to store import results and downloadable files.
"""
from odoo import fields, models


class EmployeeBulkImportLog(models.Model):
    _name = 'employee.bulk.import.log'
    _description = 'Employee Bulk Import Log'
    _order = 'create_date desc'

    name = fields.Char(string='Import Name', required=True)
    create_date = fields.Datetime(string='Imported On', readonly=True)
    create_uid = fields.Many2one('res.users', string='Imported By', readonly=True)

    imported_count = fields.Integer(string='Rows Imported', readonly=True)
    failed_count = fields.Integer(string='Rows Failed', readonly=True)
    summary = fields.Text(string='Summary', readonly=True)

    output_file = fields.Binary(string='Output Excel (Passwords)', attachment=True)
    output_filename = fields.Char(string='Output Filename')

    error_file = fields.Binary(string='Error Report Excel', attachment=True)
    error_filename = fields.Char(string='Error Filename')

    def action_download_output(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_url',
            'url': (
                f'/web/content/employee.bulk.import.log/{self.id}'
                f'/output_file/{self.output_filename}?download=true'
            ),
            'target': 'self',
        }

    def action_download_errors(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_url',
            'url': (
                f'/web/content/employee.bulk.import.log/{self.id}'
                f'/error_file/{self.error_filename}?download=true'
            ),
            'target': 'self',
        }
