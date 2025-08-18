#!/usr/bin/env python3
"""
Main processing script for PDF data extraction to Excel/CSV and SQLite database.
Handles both Annexure IV and Annexure VI document types.
"""

import os
import sys
import time
import argparse
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent
sys.path.append(str(project_root))

from data_extractors.annexure4 import AnnexureIVExtractor
from data_extractors.annexure6 import main as process_annexure6
from database import DatabaseManager
from excel_exporters import ExcelExporter, export_all_data

class DocumentProcessor:
    """Main class for processing PDF documents and exporting data"""
    
    def __init__(self):
        self.db_manager = DatabaseManager()
        self.annexure4_extractor = AnnexureIVExtractor()
        self.excel_exporter = ExcelExporter()
        
    def process_single_pdf(self, pdf_path, document_type='auto'):
        """Process a single PDF file"""
        
        if not os.path.exists(pdf_path):
            print(f"❌ Error: File not found: {pdf_path}")
            return False
        
        print(f"\n📄 Processing: {pdf_path}")
        start_time = time.time()
        
        try:
            if document_type == 'auto':
                document_type = self.detect_document_type(pdf_path)
            
            if document_type == 'annexure4':
                result = self.process_annexure4(pdf_path)
            elif document_type == 'annexure6':
                result = self.process_annexure6(pdf_path)
            else:
                print(f"❌ Unknown document type: {document_type}")
                return False
            
            processing_time = time.time() - start_time
            
            # Log the processing
            self.db_manager.log_processing(
                file_path=pdf_path,
                file_name=os.path.basename(pdf_path),
                processing_type=document_type,
                status='success',
                records_processed=1,
                processing_time=processing_time
            )
            
            print(f"✅ Successfully processed in {processing_time:.2f} seconds")
            return True
            
        except Exception as e:
            processing_time = time.time() - start_time
            error_msg = str(e)
            
            print(f"❌ Error processing {pdf_path}: {error_msg}")
            
            # Log the error
            self.db_manager.log_processing(
                file_path=pdf_path,
                file_name=os.path.basename(pdf_path),
                processing_type=document_type,
                status='error',
                error_message=error_msg,
                processing_time=processing_time
            )
            
            return False
    
    def detect_document_type(self, pdf_path):
        """Auto-detect document type based on filename or content"""
        filename = os.path.basename(pdf_path).lower()
        
        if 'annexure' in filename and ('iv' in filename or '4' in filename):
            return 'annexure4'
        elif 'annexure' in filename and ('vi' in filename or '6' in filename):
            return 'annexure6'
        elif 'declaration' in filename:
            return 'annexure4'
        elif 'invoice' in filename or 'bill' in filename:
            return 'annexure6'
        else:
            # Default to annexure4 if unsure
            print(f"⚠️  Could not auto-detect document type for {filename}, defaulting to Annexure IV")
            return 'annexure4'
    
    def process_annexure4(self, pdf_path):
        """Process Annexure IV document"""
        result = self.annexure4_extractor.process_pdf(pdf_path)
        print(f"✅ Annexure IV processed - Record ID: {result['record_id']}")
        return result
    
    def process_annexure6(self, pdf_path):
        """Process Annexure VI document"""
        # For Annexure VI, we'll use the existing script but integrate with our database
        print("🔄 Processing Annexure VI (Invoice) document...")
        
        # This would need to be modified to integrate with our database
        # For now, just call the existing function
        try:
            process_annexure6()  # This processes the hardcoded PDF
            print("✅ Annexure VI processed successfully")
            return {'status': 'success'}
        except Exception as e:
            print(f"❌ Error processing Annexure VI: {e}")
            raise
    
    def process_directory(self, directory_path, document_type='auto'):
        """Process all PDF files in a directory"""
        
        if not os.path.exists(directory_path):
            print(f"❌ Error: Directory not found: {directory_path}")
            return
        
        pdf_files = []
        for root, dirs, files in os.walk(directory_path):
            for file in files:
                if file.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(root, file))
        
        if not pdf_files:
            print(f"⚠️  No PDF files found in {directory_path}")
            return
        
        print(f"\n📁 Found {len(pdf_files)} PDF files to process")
        
        successful = 0
        failed = 0
        
        for pdf_path in pdf_files:
            if self.process_single_pdf(pdf_path, document_type):
                successful += 1
            else:
                failed += 1
        
        print(f"\n📊 Processing Summary:")
        print(f"   ✅ Successful: {successful}")
        print(f"   ❌ Failed: {failed}")
        print(f"   📄 Total: {len(pdf_files)}")
    
    def export_data(self, export_format='all'):
        """Export processed data to Excel/CSV"""
        
        print("\n📤 Exporting data...")
        
        if export_format in ['all', 'excel']:
            # Export all data formats
            files = export_all_data(self.db_manager.db_path)
            return files
        
        return []
    
    def show_summary(self):
        """Show summary of processed data"""
        
        print("\n📊 Data Summary:")
        
        # Annexure IV summary
        annexure4_summary = self.db_manager.get_annexure4_summary()
        if not annexure4_summary.empty:
            print(f"\n📋 Annexure IV Documents: {len(annexure4_summary)}")
            for _, row in annexure4_summary.head(5).iterrows():
                print(f"   • {row['file_name']} - {row['parts_count']} parts ({row['created_at'][:16]})")
            if len(annexure4_summary) > 5:
                print(f"   ... and {len(annexure4_summary) - 5} more")
        
        # Annexure VI summary
        annexure6_summary = self.db_manager.get_annexure6_summary()
        if not annexure6_summary.empty:
            print(f"\n🧾 Annexure VI Documents: {len(annexure6_summary)}")
            for _, row in annexure6_summary.head(5).iterrows():
                print(f"   • {row['file_name']} - {row['item_count']} items (₹{row['total_value']:.2f})")
            if len(annexure6_summary) > 5:
                print(f"   ... and {len(annexure6_summary) - 5} more")
        
        # Processing log
        log_summary = self.db_manager.get_processing_log()
        if not log_summary.empty:
            recent_errors = log_summary[log_summary['status'] == 'error'].head(3)
            if not recent_errors.empty:
                print(f"\n⚠️  Recent Errors:")
                for _, row in recent_errors.iterrows():
                    print(f"   • {row['file_name']}: {row['error_message'][:50]}...")

def main():
    """Main function with command line interface"""
    
    parser = argparse.ArgumentParser(description='PDF Data Extraction Tool')
    parser.add_argument('input', nargs='?', help='Input PDF file or directory')
    parser.add_argument('--type', choices=['auto', 'annexure4', 'annexure6'], 
                       default='auto', help='Document type (default: auto)')
    parser.add_argument('--export', action='store_true', 
                       help='Export data to Excel after processing')
    parser.add_argument('--summary', action='store_true', 
                       help='Show summary of processed data')
    parser.add_argument('--export-only', action='store_true',
                       help='Only export existing data without processing new files')
    
    args = parser.parse_args()
    
    # Initialize processor
    processor = DocumentProcessor()
    
    print("🚀 PDF Data Extraction Tool")
    print("=" * 50)
    
    # Handle export-only mode
    if args.export_only:
        files = processor.export_data()
        if files:
            print(f"\n✅ Exported {len(files)} files")
        return
    
    # Handle summary mode
    if args.summary:
        processor.show_summary()
        return
    
    # Process input if provided
    if args.input:
        if os.path.isfile(args.input):
            processor.process_single_pdf(args.input, args.type)
        elif os.path.isdir(args.input):
            processor.process_directory(args.input, args.type)
        else:
            print(f"❌ Error: Invalid input path: {args.input}")
            return
    else:
        # Interactive mode if no input provided
        print("\n🎯 Interactive Mode")
        print("1. Process single PDF file")
        print("2. Process directory of PDFs")
        print("3. Export existing data")
        print("4. Show data summary")
        print("5. Exit")
        
        while True:
            choice = input("\nSelect option (1-5): ").strip()
            
            if choice == '1':
                pdf_path = input("Enter PDF file path: ").strip()
                if pdf_path:
                    processor.process_single_pdf(pdf_path)
            
            elif choice == '2':
                dir_path = input("Enter directory path: ").strip()
                if dir_path:
                    processor.process_directory(dir_path)
            
            elif choice == '3':
                files = processor.export_data()
                if files:
                    print(f"\n✅ Exported {len(files)} files")
            
            elif choice == '4':
                processor.show_summary()
            
            elif choice == '5':
                print("👋 Goodbye!")
                break
            
            else:
                print("❌ Invalid choice. Please select 1-5.")
    
    # Auto-export if requested
    if args.export:
        files = processor.export_data()
        if files:
            print(f"\n✅ Exported {len(files)} files")

if __name__ == "__main__":
    main()
