import sys
import os

def main():
    # 1. Get File Path
    default_pdf = 'Point, Line, and Edge Detection.pdf'
    pdf_input = input(f"Enter PDF file path (default: {default_pdf}): ").strip()
    if not pdf_input:
        pdf_file = default_pdf
    else:
        # Remove quotes if user dragged and dropped file
        pdf_file = pdf_input.strip('"\'')
    
    if not os.path.exists(pdf_file):
        print(f"Error: File '{pdf_file}' not found.")
        return

    # 2. Get Page Limit
    limit_input = input("Enter number of pages to process (leave empty for all pages): ").strip()
    max_pages = None
    if limit_input:
        try:
            max_pages = int(limit_input)
            if max_pages <= 0:
                print("Invalid page number. Processing all pages.")
                max_pages = None
        except ValueError:
            print("Invalid input. Processing all pages.")

    print("\nSelect an option:")
    print("1. Generate English DOCX (Original Layout)")
    print("2. Generate Arabic DOCX (AI Translated)")
    
    choice = input("Enter your choice (1 or 2): ").strip()
    
    if choice == '1':
        print(f"\nRunning English conversion for '{pdf_file}'...")
        if max_pages:
             print(f"Limit: First {max_pages} pages.")
        
        try:
            import gen_docx_english
            gen_docx_english.process_all_pages(pdf_file=pdf_file, max_pages=max_pages)
            print("Done.")
        except ImportError:
            print("Error: gen_docx_english.py not found.")
        except Exception as e:
            print(f"Error running English conversion: {e}")
            
    elif choice == '2':
        print(f"\nRunning Arabic AI translation for '{pdf_file}'...")
        if max_pages:
             print(f"Limit: First {max_pages} pages.")

        try:
            import gen_docx_arabic_ai
            gen_docx_arabic_ai.process_all_pages(pdf_file=pdf_file, max_pages=max_pages)
            print("Done.")
        except ImportError:
            print("Error: gen_docx_arabic_ai.py not found.")
        except Exception as e:
            print(f"Error running Arabic translation: {e}")
            
    else:
        print("Invalid choice.")

if __name__ == "__main__":
    main()
