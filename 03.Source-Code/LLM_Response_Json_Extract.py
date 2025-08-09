import re
import os
import json
import json5

# Standardized JSON Extraction from Mixed LLM Responses

class FileProcessor:
    """Central JSON file handler"""

    @staticmethod
    def extract_json_from_content(content):
        """
        Extract JSON from mixed content.
        Supports Markdown code blocks, stand-alone JSON objects/arrays.
        """
        # Try to match Markdown JSON code block
        markdown_match = re.search(r'```json\s*(.*?)\s*```', content, re.DOTALL)
        if markdown_match:
            return markdown_match.group(1).strip()

        # Try to match stand-alone JSON object
        json_obj_match = re.search(r'\{.*\}', content, re.DOTALL)
        if json_obj_match:
            return json_obj_match.group(0).strip()

        # Try to match JSON array
        json_array_match = re.search(r'\[.*\]', content, re.DOTALL)
        if json_array_match:
            return json_array_match.group(0).strip()

        # Fallback: return entire content
        return content

    @staticmethod
    def fix_illegal_escapes(json_str):
        """Fix illegal escape sequences in JSON string"""
        illegal_escape_pattern = r'\\([^"\\/bfnrtu])'
        return re.sub(illegal_escape_pattern, r'\\\\\1', json_str)

    @staticmethod
    def structural_repair(json_str):
        """Intelligently repair JSON structure"""
        # Fix illegal escapes
        json_str = FileProcessor.fix_illegal_escapes(json_str)

        # Balance brackets and quotes
        stack = []
        in_string = False
        result = []
        i = 0

        # First pass: balance brackets and quotes
        while i < len(json_str):
            char = json_str[i]

            # Handle escape sequences
            if char == '\\' and i + 1 < len(json_str):
                result.append(char + json_str[i + 1])
                i += 2
                continue

            if char == '"':
                in_string = not in_string

            if not in_string:
                if char in ('{', '['):
                    stack.append(char)
                elif char == '}' and stack and stack[-1] == '{':
                    stack.pop()
                elif char == ']' and stack and stack[-1] == '[':
                    stack.pop()

            result.append(char)
            i += 1

        # Close unbalanced brackets
        for left in reversed(stack):
            result.append('}' if left == '{' else ']')

        # Second pass: fix truncated strings
        repaired = []
        in_string = False
        quote_count = 0
        for char in result:
            if char == '"' and (not repaired or repaired[-1] != '\\'):
                quote_count += 1
                in_string = not in_string
            repaired.append(char)

        # Complete odd number of quotes
        if quote_count % 2 != 0:
            repaired.append('"')

        # Third pass: fix trailing truncation
        repaired_str = ''.join(repaired)
        last_char = repaired_str[-1] if repaired_str else ''
        if not re.search(r'[\]}\d"truefalsenull]$', last_char):
            if re.search(r':\s*[{\[]', repaired_str):
                repaired_str += ']}' if '{' in repaired_str else ']]'
            else:
                repaired_str += '}' if '{' in repaired_str else ']'

        return repaired_str

    @staticmethod
    def remove_json_comments(json_str):
        """Remove comments from JSON"""
        pattern = r'(\s*//.*?\n)|(/\*.*?\*/)'
        return re.sub(pattern, '', json_str, flags=re.DOTALL)

    @staticmethod
    def safe_json_parse(json_str):
        """JSON parsing with auto-repair"""
        try:
            return json.loads(json_str)
        except json.JSONDecodeError:
            fixed = re.sub(r',\s*}(?=\s*})', '}', json_str)  # Remove trailing commas
            fixed = re.sub(r'[\x00-\x1F\x7F]', '', fixed)    # Strip control chars
            try:
                return json.loads(fixed)
            except:
                return json5.loads(fixed)  # Fallback to permissive json5

    @staticmethod
    def load_json(file_path):
        """Smart load JSON file"""
        # Try multiple encodings
        encodings = ['utf-8-sig', 'utf-8', 'gb18030', 'latin-1']
        content = None

        for enc in encodings:
            try:
                with open(file_path, 'r', encoding=enc) as f:
                    content = f.read()
                break
            except:
                continue

        if not content:
            raise ValueError("All encoding attempts failed")

        # Extract JSON from content
        json_content = FileProcessor.extract_json_from_content(content)

        # Remove comments
        cleaned_content = FileProcessor.remove_json_comments(json_content)

        # Structural repair
        repaired_content = FileProcessor.structural_repair(cleaned_content)

        # Parse JSON
        try:
            return FileProcessor.safe_json_parse(repaired_content)
        except Exception as e:
            # Last resort: parse entire file content
            try:
                return FileProcessor.safe_json_parse(cleaned_content)
            except:
                raise ValueError(f"JSON parsing failed: {str(e)}")

    @staticmethod
    def save_extracted_json(json_data, output_path):
        """Save extracted JSON to file"""
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, indent=2, ensure_ascii=False)

# Usage example
if __name__ == "__main__":
    # Config paths – supports both file and directory input
    INPUT_PATH = r"../Input Directory"  # File or directory
    OUTPUT_DIR = r"../Output Directory"  # Output JSON directory

    # Determine whether input is file or directory
    if os.path.isfile(INPUT_PATH):
        file_list = [INPUT_PATH]
    elif os.path.isdir(INPUT_PATH):
        file_list = []
        for root, _, files in os.walk(INPUT_PATH):
            for file in files:
                file_list.append(os.path.join(root, file))
    else:
        raise ValueError(f"Input path does not exist: {INPUT_PATH}")

    # Process all files
    for file_path in file_list:
        try:
            # Load and repair JSON
            json_data = FileProcessor.load_json(file_path)

            # Generate output path
            base_name = os.path.basename(file_path)
            file_name, file_ext = os.path.splitext(base_name)
            output_path = os.path.join(OUTPUT_DIR, f"{file_name}.json")

            # Save repaired JSON
            FileProcessor.save_extracted_json(json_data, output_path)
            print(f"Processed successfully: {file_path} → {output_path}")

        except Exception as e:
            print(f"Processing failed: {file_path} | Error: {str(e)}")