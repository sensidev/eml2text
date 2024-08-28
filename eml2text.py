import os
import glob
import shutil
import sys
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime

ROOT_PATH = os.path.abspath(os.getcwd())
DEFAULT_EML_FILES_FOLDER_PATH = os.path.join(ROOT_PATH, 'samples')
TEXT_FOLDER_PATH = os.path.join(ROOT_PATH, 'text')
OUTPUT_TEXT_FILE_PATH = os.path.join(ROOT_PATH, 'output.txt')


def extract_text_and_metadata_from_eml(eml_file):
    """Extracts all headers, date, and plain text from an .eml file, logging a warning if attachments are ignored."""
    with open(eml_file, 'rb') as file:
        msg = BytesParser(policy=policy.default).parse(file)

    # Extract metadata (headers)
    headers = {
        'From': msg['From'],
        'To': msg['To'],
        'Subject': msg['Subject'],
        'Date': msg['Date']
    }

    # Parse the date header to a datetime object
    email_date = parsedate_to_datetime(headers['Date'])

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

    return email_date, text_content


def convert_eml_to_text_with_metadata(eml_folder, output_folder):
    """
    Converts all *.eml files in a folder to individual text files,
    including metadata, and returns a sorted list of files by date.
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    eml_files = glob.glob(os.path.join(eml_folder, '*.eml'))

    email_data = []

    for eml_file in eml_files:
        email_date, text_content = extract_text_and_metadata_from_eml(eml_file)
        base_name = os.path.basename(eml_file)
        txt_file_name = os.path.splitext(base_name)[0] + '.txt'
        txt_file_path = os.path.join(output_folder, txt_file_name)

        with open(txt_file_path, 'w', encoding='utf-8') as text_file:
            text_file.write(text_content)

        # Append the extracted data to the list
        email_data.append((email_date, txt_file_path))

    # Sort the emails by date
    email_data.sort(key=lambda x: x[0])

    return email_data


def merge_text_files(sorted_email_data, merged_output_file):
    """Merges all text files in a folder into one big text file with start and end markers, sorted chronologically."""
    with open(merged_output_file, 'w', encoding='utf-8') as outfile:
        for email_date, txt_file in sorted_email_data:
            file_name = os.path.basename(txt_file)

            # Write start marker with file name
            outfile.write(f"--- START OF {file_name} ({email_date.strftime('%Y-%m-%d %H:%M:%S')}) ---\n")

            with open(txt_file, 'r', encoding='utf-8') as infile:
                outfile.write(infile.read())

            # Write end marker with file name
            outfile.write(f"\n--- END OF {file_name} ---\n\n")
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

    sorted_email_data = convert_eml_to_text_with_metadata(eml_folder, output_folder)
    eml_file_count = len(sorted_email_data)

    merge_text_files(sorted_email_data, merged_output_file)

    print(f'DONE! All {eml_file_count} .eml files have been converted,'
          f' sorted chronologically and merged into {merged_output_file}')


if __name__ == "__main__":
    main()
