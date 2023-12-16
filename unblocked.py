import os

def unblock_file(file_path):
    zone_identifier_stream = file_path + ":Zone.Identifier"
    if os.path.exists(zone_identifier_stream):
        try:
            os.remove(zone_identifier_stream)
            print(f"Unblocked file: {file_path}")
        except Exception as e:
            print(f"Error unblocking file: {e}")
    else:
        print(f"No Zone.Identifier stream found for {file_path}")

# Example usage
file_path = 'E:/workspace/scraping/print-cell-of-excel-file/albury_R6andR7_20231216.xlsx'
unblock_file(file_path)