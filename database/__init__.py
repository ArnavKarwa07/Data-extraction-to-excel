import sqlite3
import pandas as pd
import os
from datetime import datetime

class DatabaseManager:
    """Handles all database operations for the data extraction project"""
    
    def __init__(self, db_path="database/extraction_data.db"):
        self.db_path = db_path
        self.ensure_database_exists()
    
    def ensure_database_exists(self):
        """Create database and all necessary tables"""
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Annexure IV main table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS annexure4_data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_count INTEGER,
                pli_request_no TEXT,
                ifci_no TEXT,
                file_name TEXT,
                date_processed TEXT,
                applicant TEXT,
                supplier TEXT,
                subject TEXT,
                currency_name TEXT,
                reference_date TEXT,
                foreign_exchange_rate TEXT,
                iec_code TEXT,
                sr_no_1 TEXT,
                sr_no_2 TEXT,
                sr_no_3 TEXT,
                sr_no_4 TEXT,
                sr_no_5 TEXT,
                sr_no_6 TEXT,
                applicant_signatory_ca_fa TEXT,
                udin_values TEXT,
                udin_values_2 TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Annexure IV parts/components table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS annexure4_parts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                annexure4_id INTEGER,
                part_no TEXT,
                part_description TEXT,
                selling_price_inr TEXT,
                value_direct_import_inr TEXT,
                broad_description_parts_imported TEXT,
                value_parts_imported_suppliers_inr TEXT,
                broad_description_parts_imported_suppliers TEXT,
                local_content TEXT,
                dva_percentage TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (annexure4_id) REFERENCES annexure4_data (id)
            )
        ''')
        
        # Annexure VI main table (for invoice data)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS annexure6_data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pli_request_no TEXT,
                ifci_no TEXT,
                file_name TEXT,
                invoice_issued_to TEXT,
                invoice_issued_to_gstin TEXT,
                serial_no INTEGER,
                irn_no TEXT,
                invoice_no TEXT,
                date_processed TEXT,
                local_supplier_name TEXT,
                local_supplier_gstin TEXT,
                part_component_name TEXT,
                hsn_code TEXT,
                value_net_gst REAL,
                quantity REAL,
                value_per_piece REAL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Processing log table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS processing_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_path TEXT,
                file_name TEXT,
                processing_type TEXT,
                status TEXT,
                error_message TEXT,
                records_processed INTEGER,
                processing_time_seconds REAL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def insert_annexure4_data(self, header_info, parts_data, file_name):
        """Insert Annexure IV data into database"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            # Insert main record
            cursor.execute('''
                INSERT INTO annexure4_data (
                    file_count, pli_request_no, ifci_no, file_name, date_processed,
                    applicant, supplier, subject, currency_name, reference_date,
                    foreign_exchange_rate, iec_code, sr_no_1, sr_no_2, sr_no_3,
                    sr_no_4, sr_no_5, sr_no_6, applicant_signatory_ca_fa,
                    udin_values, udin_values_2, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                1,  # file_count
                header_info.get('pli_request_no', ''),
                header_info.get('ifci_no', ''),
                file_name,
                header_info.get('date', ''),
                header_info.get('applicant', ''),
                header_info.get('supplier', ''),
                header_info.get('subject', ''),
                header_info.get('currency', ''),
                header_info.get('reference_date', ''),
                header_info.get('foreign_exchange_rate', ''),
                header_info.get('iec_code', ''),
                header_info.get('sr_no_1', ''),
                header_info.get('sr_no_2', ''),
                header_info.get('sr_no_3', ''),
                header_info.get('sr_no_4', ''),
                header_info.get('sr_no_5', ''),
                header_info.get('sr_no_6', ''),
                header_info.get('applicant_signatory_ca_fa', ''),
                header_info.get('udin_values', ''),
                header_info.get('udin_values_2', ''),
                datetime.now().isoformat()
            ))
            
            annexure4_id = cursor.lastrowid
            
            # Insert parts data
            for part in parts_data:
                cursor.execute('''
                    INSERT INTO annexure4_parts (
                        annexure4_id, part_no, part_description, selling_price_inr,
                        value_direct_import_inr, broad_description_parts_imported,
                        value_parts_imported_suppliers_inr, broad_description_parts_imported_suppliers,
                        local_content, dva_percentage
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    annexure4_id,
                    part.get('part_no', ''),
                    part.get('part_description', ''),
                    part.get('selling_price_inr', ''),
                    part.get('value_direct_import_inr', ''),
                    part.get('broad_description_parts_imported', ''),
                    part.get('value_parts_imported_suppliers_inr', ''),
                    part.get('broad_description_parts_imported_suppliers', ''),
                    part.get('local_content', ''),
                    part.get('dva_percentage', '')
                ))
            
            conn.commit()
            return annexure4_id
            
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()
    
    def insert_annexure6_data(self, data_rows):
        """Insert Annexure VI (invoice) data into database"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            for row in data_rows:
                cursor.execute('''
                    INSERT INTO annexure6_data (
                        pli_request_no, ifci_no, file_name, invoice_issued_to,
                        invoice_issued_to_gstin, serial_no, irn_no, invoice_no,
                        date_processed, local_supplier_name, local_supplier_gstin,
                        part_component_name, hsn_code, value_net_gst, quantity,
                        value_per_piece, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    row.get('PLI Request No', ''),
                    row.get('IFCI No', ''),
                    row.get('File Name', ''),
                    row.get('invoice issued to', ''),
                    row.get('invoice issued to GSTIN', ''),
                    row.get('#', 0),
                    row.get('IRN#', ''),
                    row.get('Invoice#', ''),
                    row.get('Date', ''),
                    row.get('Name of Local Supplier', ''),
                    row.get('GSTIN of Local Supplier', ''),
                    row.get('Name of Part/Component', ''),
                    row.get('HSN Code', ''),
                    row.get('Value (net of GST) (Rs.)', 0),
                    row.get('Quantity', 0),
                    row.get('Value per piece (net of GST) (Rs.)', 0),
                    datetime.now().isoformat()
                ))
            
            conn.commit()
            return len(data_rows)
            
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()
    
    def log_processing(self, file_path, file_name, processing_type, status, 
                      error_message=None, records_processed=0, processing_time=0):
        """Log processing results"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            INSERT INTO processing_log (
                file_path, file_name, processing_type, status, error_message,
                records_processed, processing_time_seconds
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (file_path, file_name, processing_type, status, error_message, 
              records_processed, processing_time))
        
        conn.commit()
        conn.close()
    
    def get_annexure4_summary(self):
        """Get summary of Annexure IV data"""
        conn = sqlite3.connect(self.db_path)
        
        query = '''
            SELECT 
                a.id,
                a.file_name,
                a.date_processed,
                a.applicant,
                a.supplier,
                COUNT(p.id) as parts_count,
                a.created_at
            FROM annexure4_data a
            LEFT JOIN annexure4_parts p ON a.id = p.annexure4_id
            GROUP BY a.id
            ORDER BY a.created_at DESC
        '''
        
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df
    
    def get_annexure6_summary(self):
        """Get summary of Annexure VI data"""
        conn = sqlite3.connect(self.db_path)
        
        query = '''
            SELECT 
                file_name,
                COUNT(*) as item_count,
                SUM(value_net_gst) as total_value,
                MIN(date_processed) as earliest_date,
                MAX(date_processed) as latest_date,
                COUNT(DISTINCT local_supplier_name) as supplier_count
            FROM annexure6_data
            GROUP BY file_name
            ORDER BY MAX(created_at) DESC
        '''
        
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df
    
    def export_annexure4_to_excel(self, output_file='Annexure4_Export.xlsx'):
        """Export Annexure IV data to Excel in the required format"""
        conn = sqlite3.connect(self.db_path)
        
        # Main query for detailed data
        main_query = '''
            SELECT 
                a.file_count,
                a.pli_request_no,
                a.ifci_no,
                a.file_name,
                a.date_processed,
                a.applicant,
                a.supplier,
                a.subject,
                a.currency_name,
                a.reference_date,
                a.foreign_exchange_rate,
                a.iec_code,
                p.part_no,
                p.part_description,
                p.selling_price_inr,
                p.value_direct_import_inr,
                p.broad_description_parts_imported,
                p.value_parts_imported_suppliers_inr,
                p.broad_description_parts_imported_suppliers,
                p.local_content,
                p.dva_percentage
            FROM annexure4_data a
            LEFT JOIN annexure4_parts p ON a.id = p.annexure4_id
            ORDER BY a.id, p.id
        '''
        
        # Summary query (like image 4)
        summary_query = '''
            SELECT 
                a.file_count as "File Count",
                a.pli_request_no as "PLI Request No",
                a.ifci_no as "IFCI No",
                a.file_name as "File Name",
                a.date_processed as "Date",
                a.file_name as "Annexure IV - Declaration from PDF",
                a.applicant as "Applicant",
                a.supplier as "Supplier",
                a.subject as "Subject",
                a.sr_no_1 as "Sr No 1",
                a.sr_no_2 as "Sr No 2",
                a.sr_no_3 as "Sr No 3",
                a.sr_no_4 as "Sr No 4",
                a.sr_no_5 as "Sr No 5",
                a.sr_no_6 as "Sr No 6",
                a.applicant_signatory_ca_fa as "Applicant signatory CA / SA",
                a.udin_values as "UDIN values",
                a.udin_values_2 as "UDIN values 2",
                GROUP_CONCAT(
                    p.part_no || ' ' || p.part_description || ' ' || 
                    p.selling_price_inr || ' ' || p.value_direct_import_inr || ' ' ||
                    p.broad_description_parts_imported || ' ' || p.value_parts_imported_suppliers_inr || ' ' ||
                    p.broad_description_parts_imported_suppliers || ' ' || p.local_content || ' ' || p.dva_percentage,
                    ' | '
                ) as "Content text matching exact as per output format"
            FROM annexure4_data a
            LEFT JOIN annexure4_parts p ON a.id = p.annexure4_id
            GROUP BY a.id
            ORDER BY a.id
        '''
        
        main_df = pd.read_sql_query(main_query, conn)
        summary_df = pd.read_sql_query(summary_query, conn)
        
        conn.close()
        
        # Create Excel file with multiple sheets
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Detailed data
            main_df.to_excel(writer, sheet_name='Detailed Data', index=False)
            
            # Summary format (like image 4)
            summary_df.to_excel(writer, sheet_name='Annexure IV Summary', index=False)
            
            # Statistics
            stats_df = pd.DataFrame({
                'Metric': ['Total Files Processed', 'Total Parts/Components', 'Processing Date'],
                'Value': [
                    len(summary_df),
                    len(main_df[main_df['part_no'].notna()]),
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ]
            })
            stats_df.to_excel(writer, sheet_name='Statistics', index=False)
        
        print(f"✓ Annexure IV data exported to {output_file}")
        return output_file
    
    def export_annexure6_to_excel(self, output_file='Annexure6_Export.xlsx'):
        """Export Annexure VI data to Excel"""
        conn = sqlite3.connect(self.db_path)
        
        query = '''
            SELECT * FROM annexure6_data
            ORDER BY created_at DESC
        '''
        
        df = pd.read_sql_query(query, conn)
        conn.close()
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Invoice Data', index=False)
        
        print(f"✓ Annexure VI data exported to {output_file}")
        return output_file
    
    def get_processing_log(self):
        """Get processing log"""
        conn = sqlite3.connect(self.db_path)
        
        query = '''
            SELECT * FROM processing_log
            ORDER BY created_at DESC
        '''
        
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df
    
    def clear_all_data(self, confirm=False):
        """Clear all data from database (use with caution)"""
        if not confirm:
            print("Warning: This will delete all data. Call with confirm=True to proceed.")
            return
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        tables = ['annexure4_parts', 'annexure4_data', 'annexure6_data', 'processing_log']
        
        for table in tables:
            cursor.execute(f'DELETE FROM {table}')
        
        conn.commit()
        conn.close()
        print("All data cleared from database.")
