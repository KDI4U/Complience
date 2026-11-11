ESIC AUTO - Project Structure

1. Put PostgreSQL credentials in .env
2. Edit modules/db_download.py and replace SQL_QUERY.
3. Run:
       python modules/db_download.py

The fresh DB export will be created at:
       data/db/DB_Data.xlsx

Later:
- esic_download.py = ESIC portal download
- excel_merge.py   = merge ESIC files
- compare_data.py  = DB vs ESIC comparison
- hrms_upload.py   = upload new records to HRMS
- main.py          = run the complete automation
