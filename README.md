# 

<div align="center">
   <h1>python_ocr_pdf_to_excel</h1>
</div>

### Input Images :

<div align="center">
   <img src=https://github.com/LucaIT523/python_ocr_pdf_to_excel/blob/main/image/1.png>
</div>





<div align="center">
   <img src=https://github.com/LucaIT523/python_ocr_pdf_to_excel/blob/main/image/2.png>
</div>



### 1. Core Architecture

```
# Core Components
PaddleOCR(use_angle_cls=True, lang='en')  # AI-powered OCR engine
xlsxwriter.Workbook()                     # Excel report generator
cv2.imread()/cv2.imwrite()                # Image processing
pdf2image (implied)                       # PDF-to-image conversion
```

#### Key Features:

- **Document Processing Pipeline**: `PDF → Images → Text Extraction → Data Structuring → Excel Report`
- **Layout Analysis**: Uses coordinate-based region cropping
- **Multi-period Comparison**: Handles N and N-1 fiscal years

### 2. Workflow Breakdown

#### 2.1 Document Preparation

```
# Implied PDF processing (via pdf2image)
for img_idx in range(5):  # Processes 5 pages
    img_path = './image/'+str(img_idx+1)+'.png'
```

#### Document Handling:

- Converts PDF to PNG images (1 image per page)
- Processes first 5 pages of document
- Requires pre-processed images in `/image` folder

#### 2.2 Key Data Extraction

```
pythonCopy# Core OCR Process
result = pdocr.ocr(img_path, cls=False)

# Specialized Field Detection
check_key()       # Identifies 8-digit account numbers (>100000)
compare_key()     # Finds "Exercice N/N-1" headers
```

#### Pattern Matching:

- **Account Numbers**: 8-digit numeric values
- **Fiscal Periods**: "Exercice N" (Current) and "Exercice N-1" (Previous)
- **Header Validation**: Exact string matching for financial periods

### 3. Image Processing

#### 3.1 Region Cropping

```
sub_image_proc()  # Creates focused analysis regions:
                  # - Account number rows
                  # - Financial period columns
```

#### Coordinate Logic:

- **Vertical Cropping**: `mainkey_list[idx][0][1] - 5` to `mainkey_list[idx][2][1] + 5`
- **Horizontal Expansion**: `w_e = dateInfo_list[1][2][0] + 50`
- **Position Offsets**: Compensates for OCR coordinate variations

#### 3.2 Numerical Extraction

```
get_num_info()  # Cleans OCR results:
                # 1. Removes hyphens/spaces
                # 2. Converts to float
                # 3. Handles empty values
```

#### Data Sanitization:

- `num_data.replace('-', '')`: Removes negative signs (potential issue)
- `strip()`: Eliminates padding spaces
- Fallback to `0.0` for empty values

### 4. Report Generation

#### Excel Structure:

```
worksheet.write('A1', 'Account Number')  # Column A
worksheet.write('B1', 'Account Title')   # Column B
worksheet.write('C1', 'Date')            # Column C
worksheet.write('D1', 'Amount')         # Column D
```

#### Data Organization:

- **Row Duplication**: Each account creates 2 rows (N and N-1)
- **Date Formatting**: `str_data[:10]` truncates dates to 10 characters
- **Progressive Indexing**: `row += 2` maintains row spacing

### 5. Key Technical Considerations

#### 5.1 OCR Limitations

- **Font Sensitivity**: Depends on PaddleOCR's English model accuracy
- **Layout Dependency**: Requires consistent document formatting
- **Coordinate Drift**: Fixed pixel offsets (`+35`, `+20`) may fail with layout changes

#### 5.2 Financial Data Handling

- **Negative Values**: Current implementation removes hyphens (may corrupt negatives)
- **Decimal Handling**: No explicit decimal point validation
- **Currency Symbols**: Not preserved in current implementation







### **Contact Us**

For any inquiries or questions, please contact us.

telegram : @topdev1012

email :  skymorning523@gmail.com

Teams :  https://teams.live.com/l/invite/FEA2FDDFSy11sfuegI