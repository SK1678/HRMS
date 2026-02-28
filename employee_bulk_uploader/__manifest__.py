{
    'name': 'Employee Bulk Uploader',
    'version': '19.0.1.0.0',
    'category': 'Human Resources',
    'summary': 'Bulk import employees from Excel with preview, validation, and user auto-creation.',
    'description': """
        Upload an Excel (.xlsx) file to bulk-create/update hr.employee records.
        Features:
        - File upload & parse
        - Row-by-row validation with preview table (no DB write until confirmed)
        - Auto-creates res.users per employee (login = work_email, random password)
        - Post-import Excel output with passwords
        - Partial import with savepoints when stop_on_error = False
    """,
    'author': 'Betopia Group -Meher Kanti Sarkar',
    'depends': [
        'base',
        'hr',
        'beto_hr',
    ],
    'data': [
        'security/security.xml',
        'security/ir.model.access.csv',
        'wizard/employee_bulk_upload_views.xml',
        'models/import_log_views.xml',
        'data/menu.xml',
    ],
    'external_dependencies': {
        'python': ['openpyxl'],
    },
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}
