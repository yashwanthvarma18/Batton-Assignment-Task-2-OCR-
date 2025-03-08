import os
import json
from openai import OpenAI
from paddleocr import PaddleOCR
import pandas as pd
from dotenv import load_dotenv

# Load environment variables ( OpenAPI key)
load_dotenv()

# Function to extract text from image and structure it as a table
def extract_structured_data(image_path):
    # Initialize OCR to extract text from image Using PaddleOCR 
    # Supported Languages for PaddleOCR (lang=): 'ch' (Chinese Simplified), 'en' (English), 'korean' (Korean),
    #'japan' (Japanese), 'chinese_cht' (Chinese Traditional), 'ta' (Tamil), 'te' (Telugu), 'ka' (Kannada), 
    #'latin' (Latin-based languages like French, Spanish, etc.), 'arabic' (Arabic), 'cyrillic'
    #(Cyrillic-based languages like Russian, Ukrainian, etc.), 'devanagari' (Devanagari script like Hindi, Marathi, etc.)

    ocr = PaddleOCR(lang='en', use_angle_cls=True, show_log=False)
    result = ocr.ocr(image_path, cls=True)

    # Store text with its position in the image
    text_elements = []
    for line in result[0]:
        text = line[1][0]  # Extract the text
        bbox = line[0]  # Bounding box (position)
        top_left_y = min(bbox[0][1], bbox[1][1], bbox[2][1], bbox[3][1])  # Top position of text
        text_elements.append((top_left_y, bbox[0][0], text))

    # Sort text by row (y-axis) and then by column (x-axis)
    sorted_elements = sorted(text_elements, key=lambda x: (x[0], x[1]))

    # Group text by rows based on vertical position
    rows = []
    current_row = []
    prev_y = None
    y_threshold = 20  # Maximum vertical gap to consider the same row (You may need to adjust this value)

    for element in sorted_elements:
        y, x, text = element
        # Check if the current text is in the same row as previous
        if prev_y is None or abs(y - prev_y) <= y_threshold:
            current_row.append((x, text))
        else:
            rows.append(current_row)
            current_row = [(x, text)]
        prev_y = y

    # Append the last row if exists
    if current_row:
        rows.append(current_row)

    # Sort each row's text by horizontal position (x-axis)
    table_data = []
    for row in rows:
        sorted_columns = sorted(row, key=lambda x: x[0])
        table_data.append([col[1] for col in sorted_columns])  # Extract text only

    return table_data

# Function to improve the extracted table using GPT model
def refine_with_chatgpt(table_data):
    # Connect to GPT model with Azure API key
    client = OpenAI(
        base_url="https://models.inference.ai.azure.com",
        api_key=os.environ["OPENAI_API_KEY"], 
    )

    # Prompt for GPT to clean and fix the extracted table
    prompt = f"""Analyze and refine this table data extracted from an image. 
Ensure:
1. The first row contains column headers.
2. Data types in each column are consistent.
3. Any split values that belong together (including formulas, special characters, etc.) are properly merged.
4. Return valid JSON format (an array of arrays).

Return ONLY the JSON with this structure:
[
  ["Header1", "Header2", ...],
  ["Value1", "Value2", ...],
  ...
]

Raw data:
{json.dumps(table_data, indent=2)}"""

    # Send the request to GPT model
    response = client.chat.completions.create(
        messages=[
            {
                "role": "system",
                "content": "You are an expert data analyst that fixes OCR-extracted tables. Return only valid JSON."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        model="gpt-4o",
        temperature=0.2,
        max_tokens=4096,
        top_p=1
    )

    # Try to extract JSON response from GPT and convert it to Python list
    try:
        json_str = response.choices[0].message.content
        json_str = json_str.replace('```json', '').replace('```', '').strip()  # Remove unwanted markdown
        return json.loads(json_str)
    except Exception as e:
        print(f"ChatGPT refinement failed: {e}")
        return table_data  # Return original table if GPT fails

# Function to save the table as Excel file
def save_to_excel(data, filename="output_table.xlsx"):
    if len(data) < 1:
        print("No data to save!")
        return

    try:
        # Create DataFrame (first row as header)
        df = pd.DataFrame(data[1:], columns=data[0])
        try:
            # xlsxwriter for better formatting
            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Sheet1', index=False)
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']

                # Adjust column width and enable text wrapping
                wrap_format = workbook.add_format({'text_wrap': True})
                for idx, col in enumerate(df.columns):
                    max_len = max(
                        df[col].astype(str).map(len).max(),
                        len(col)
                    ) + 2  # Add extra space for readability
                    worksheet.set_column(idx, idx, max_len, wrap_format)
            print(f"Successfully saved to {os.path.abspath(filename)}")
        except ModuleNotFoundError:
            # Fallback to openpyxl if xlsxwriter is not installed
            print("Module 'xlsxwriter' not found. Falling back to openpyxl (without auto-adjusted column widths or text wrap).")
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Sheet1', index=False)
            print(f"Successfully saved to {os.path.abspath(filename)} (using openpyxl)")
    except Exception as e:
        print(f"Excel save failed: {e}")

if __name__ == "__main__":
    image_path = "table3.jpg" # Path to the image file in the directory
    # Error handling
    try:
        print("Starting OCR processing...")
        raw_table = extract_structured_data(image_path)
        print("OCR extraction completed")
        
        print("Starting AI refinement...")
        refined_table = refine_with_chatgpt(raw_table)
        print("AI refinement completed")
        
        save_to_excel(refined_table)
    except Exception as e:
        print(f"Critical error: {e}")
