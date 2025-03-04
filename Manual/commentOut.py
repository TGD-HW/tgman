import os
import re

def comment_out_html_tables(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()

    # Regex pattern to find HTML tables
    pattern = re.compile(r'(<style type="text/css">.*?</table>)', re.DOTALL)

    # Replace matched patterns with commented versions
    commented_content = pattern.sub(r'<!-- \1 -->', content)

    # Write the modified content back to the file
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(commented_content)

def main():
    max_depth = 6
    base_path = os.getcwd()

    for root, _, files in os.walk(base_path):
        # Calculate the current depth
        depth = len(os.path.relpath(root, base_path).split(os.sep))

        if depth > max_depth:
            continue

        for file in files:
            if file in ['dimension.md', 'dimension.en.md']:
                file_path = os.path.join(root, file)
                comment_out_html_tables(file_path)
                print(f'Processed {file_path}')

if __name__ == '__main__':
    main()
