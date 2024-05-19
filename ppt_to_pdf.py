import os
import comtypes.client
from tqdm import tqdm
import shutil

def convert_to_pdf(input_dir, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    ppt_app = comtypes.client.CreateObject("PowerPoint.Application")
    ppt_app.Visible = True
    word_app = comtypes.client.CreateObject("Word.Application")
    word_app.Visible = True
    files = [file for root, _, files in os.walk(input_dir) for file in files]

    try:
        for file in tqdm(files, desc="Converting files"):
            input_path = os.path.join(input_dir, file)
            output_path = os.path.join(output_dir, os.path.splitext(file)[0] + ".pdf")

            if file.endswith(('.ppt', '.pptx')):
                try:
                    ppt = ppt_app.Presentations.Open(input_path)
                    ppt.SaveAs(output_path, FileFormat=32)
                    ppt.Close()
                except Exception as e:
                    print(f"Failed to convert {input_path}: {e}")

            elif file.endswith(('.doc', '.docx')):
                try:
                    doc = word_app.Documents.Open(input_path)
                    doc.SaveAs(output_path, FileFormat=17)
                    doc.Close()
                except Exception as e:
                    print(f"Failed to convert {input_path}: {e}")

            print(f"Converted {input_path} to {output_path}")
    finally:
        ppt_app.Quit()
        word_app.Quit()

def copy_existing_pdf(input_dir, output_dir):
    for root, _, files in os.walk(input_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                input_path = os.path.join(root, file)
                output_path = os.path.join(output_dir, file)

                if os.path.abspath(input_path) == os.path.abspath(output_path):
                    print(f"Skipping {input_path} as it is the same as {output_path}")
                    continue

                shutil.copy(input_path, output_path)
                print(f"Copied {input_path} to {output_path}")

def main():
    while True:
        input_directory = input("Enter the input directory path: ")
        if os.path.exists(input_directory):
            break
        print("Input directory does not exist. Please try again.")
    
    output_directory = os.path.join(input_directory, "converted_pdf")
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
        print("Output directory created successfully.")
    else:
        print("Output directory already exists.")
            
    convert_to_pdf(input_directory, output_directory)
    copy_existing_pdf(input_directory, output_directory)

if __name__ == "__main__":
    main()
