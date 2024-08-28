import os
import glob
import shutil
import sys
from email import policy
from email.parser import BytesParser

ROOT_PATH = os.path.abspath(os.getcwd())
DEFAULT_EML_FILES_FOLDER_PATH = os.path.join(ROOT_PATH, 'samples')
TEXT_FOLDER_PATH = os.path.join(ROOT_PATH, 'text')
OUTPUT_TEXT_FILE_PATH = os.path.join(ROOT_PATH, 'output.txt')


def extract_text_and_metadata_from_eml(eml_file):
    """Extracts all headers and plain text from an .eml file."""
    with open(eml_file, 'rb') as file:
        msg = BytesParser(policy=policy.default).parse(file)

    # Extract metadata (headers)
    headers = {
        'From': msg['From'],
        'To': msg['To'],
        'Subject': msg['Subject'],
        'Date': msg['Date']
    }

    # Initialize text content with metadata
    text_content = "\n".join(f"{key}: {value}" for key, value in headers.items()) + "\n\n"

    # Extract the email's body
    attachment_found = False
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == 'text/plain':
                text_content += part.get_payload(decode=True).decode(part.get_content_charset('utf-8'))
            elif part.get_content_disposition() == 'attachment':
                attachment_found = True
    else:
        text_content += msg.get_payload(decode=True).decode(msg.get_content_charset('utf-8'))

    # Log warning if attachment was found and ignored
    if attachment_found:
        print(f"Warning: The email '{os.path.basename(eml_file)}' contains one or more attachments that were ignored.")

    return text_content


def convert_eml_to_text_with_metadata(eml_folder, output_folder):
    """Converts all .eml files in a folder to individual text files, including metadata."""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    eml_files = glob.glob(os.path.join(eml_folder, '*.eml'))

    for eml_file in eml_files:
        text_content = extract_text_and_metadata_from_eml(eml_file)
        base_name = os.path.basename(eml_file)
        txt_file_name = os.path.splitext(base_name)[0] + '.txt'
        txt_file_path = os.path.join(output_folder, txt_file_name)

        with open(txt_file_path, 'w', encoding='utf-8') as text_file:
            text_file.write(text_content)


def merge_text_files(output_folder, merged_output_file):
    """Merges all text files in a folder into one big text file."""
    with open(merged_output_file, 'w', encoding='utf-8') as outfile:
        for txt_file in glob.glob(os.path.join(output_folder, '*.txt')):
            file_name = os.path.basename(txt_file)
            outfile.write(f'\n--- START: {file_name} ---\n')

            with open(txt_file, 'r', encoding='utf-8') as infile:
                outfile.write(infile.read())

            outfile.write(f'\n--- END: {file_name} ---\n\n')
            print(f'Processed {txt_file}')


def prompt_to_delete_output_folder(output_folder):
    """Prompts the user to decide whether to delete the output folder."""
    if os.path.exists(output_folder):
        user_input = input(
            f"The output folder '{output_folder}' already exists. Do you want to delete it and start fresh? (yes/no): "
        ).strip().lower()

        if user_input in ('yes', 'y'):
            shutil.rmtree(output_folder)
            print(f"Deleted the folder '{output_folder}'. Starting fresh.")
        else:
            print("Continuing without deleting the folder. Existing files may be overwritten or included in the merge.")


def main():
    if len(sys.argv) < 2:
        eml_folder = DEFAULT_EML_FILES_FOLDER_PATH
    else:
        eml_folder = sys.argv[1]

    output_folder = TEXT_FOLDER_PATH
    merged_output_file = OUTPUT_TEXT_FILE_PATH

    prompt_to_delete_output_folder(output_folder)

    convert_eml_to_text_with_metadata(eml_folder, output_folder)
    merge_text_files(output_folder, merged_output_file)
    print(f'DONE! All .eml files have been converted and merged into {merged_output_file}')


if __name__ == "__main__":
    main()
