# PowerPoint Template Merge Field Updates Required

## Current vs Required Merge Fields Analysis

### Page 1 - Title Page

**Current:** `{{Trademark Search Report}}`  
**Required:** No change needed (static text)

### Page 4 - Report Overview Order Details

**Current Fields → Required Updates:**

| Current Field                | Required Replacement                   | Excel Source                      |
| ---------------------------- | -------------------------------------- | --------------------------------- |
| `{{Contact Name}}`           | `{{client_info.client_name}}`          | Order Form > Client Info          |
| `{{Last Search date}}`       | `{{report_date}}`                      | Multiple sheets > Report Date     |
| `{{All time/Last 30 days}}`  | Static text or new field needed        | Not in Excel                      |
| `{{Word/Image/Combination}}` | `{{client_info.search_type}}`          | Order Form > Word Or Image        |
| `{{numbers}}`                | `{{client_info.gs_classes}}`           | Order Form > G&S Classes          |
| `{{SIC Code}}`               | `{{client_info.sic_code}}`             | Order Form > SIC                  |
| `{{Business Description}}`   | `{{client_info.business_nature}}`      | Order Form > Nature of business   |
| `{{UK}}`                     | `{{client_info.designated_countries}}` | Order Form > Designated Countries |

### Page 5 - Search Criteria

**Word Search Fields:**

| Current Field       | Required Replacement                             | Excel Source                       |
| ------------------- | ------------------------------------------------ | ---------------------------------- |
| `{{Exact Match}}`   | `{{word_search.0.remarks}}` or `Found/Not Found` | Order Form > Word Search > Remarks |
| `{{Similar Match}}` | `{{word_search.1.remarks}}`                      | Order Form > Word Search > Remarks |
| `{{Starts With}}`   | `{{word_search.2.remarks}}`                      | Order Form > Word Search > Remarks |
| `{{Contains}}`      | `{{word_search.3.remarks}}`                      | Order Form > Word Search > Remarks |

**Logo Search Fields:**

| Current Field                            | Required Replacement                      | Excel Source                                                 |
| ---------------------------------------- | ----------------------------------------- | ------------------------------------------------------------ |
| `{{JPEG Logo}}`                          | `{{image_search.0.image}}`                | Order Form > Image Search > Image                            |
| `{{Search Criteria}}`                    | `{{image_search.0.search_criteria}}`      | Order Form > Image Search > Search Criteria                  |
| `{{Image Class, Division, Subdivision}}` | `{{image_search.0.image_classification}}` | Order Form > Image Search > Image Class.Division.Subdivision |

### Page 6 - Classes and Terms

**Current Fields → Required Updates:**

| Current Field | Required Replacement         | Excel Source                |
| ------------- | ---------------------------- | --------------------------- |
| `{{Classes}}` | `{{client_info.gs_classes}}` | Order Form > G&S Classes    |
| `{{Terms}}`   | New field needed             | Not clearly mapped in Excel |

### Page 7 - Google Search Results

**Table structure needs complete revision:**

**Current:** Simple table with Keywords/Images, Results, Links  
**Required:** Need to map to actual Google sheet data (which appears to be missing from current Excel structure)

_Note: Google sheet was excluded from configuration as requested_

### Page 8 - Companies House UK Search Results

**Current:** Static example data  
**Required:** Dynamic table from Companies sheet

```
{{#companies}}
- Company: {{registered_companies}}
- Link: {{link}}
- Status: {{status}}
- Registration Number: {{registration_number}}
- SIC: {{sic}}
- Remarks: {{remarks}}
{{/companies}}
```

### Page 9 - Domain Name Search Results

**Current:** Simple summary table  
**Required:** Dynamic table from Domains sheet

```
{{#domains}}
- Keywords: {{keywords}}
- .com: {{dot_com}}
- .net: {{dot_net}}
- .co.uk: {{dot_co_uk}}
- .co: {{dot_co}}
- .uk: {{dot_uk}}
{{/domains}}
```

### Page 10 - Social Media Search Results

**Current:** Static table layout  
**Required:** Dynamic table from Social sheet

```
{{#social_media}}
- Keywords: {{keywords}}
- Facebook: {{facebook}}
- Instagram: {{instagram}}
- LinkedIn: {{linkedin}}
- TikTok: {{tiktok}}
- YouTube: {{youtube}}
- X: {{x_twitter}}
{{/social_media}}
```

### Pages 11+ - Trademark Registry Search Results

**Current:** Static trademark data  
**Required:** Dynamic data from Trademarks sheet

```
{{#trademarks}}
- Office: {{office}}
- App Number: {{app_number}}
- Status: {{status}}
- Mark Type: {{mark_type}}
- Mark Text: {{mark_text}}
- Filing Date: {{filing_date}}
- Classes: {{classes}}
- Owner Name: {{owner_name}}
- Owner Location: {{owner_location}}
- Industry: {{industry}}
- Primary Goods/Services: {{primary_goods_services}}
- Representative: {{representative}}
{{/trademarks}}
```

### Pages 12-28 - Risk Analysis & Recommendations

**Current:** Placeholder fields like `{{Trademark C2}}`, `{{Trademark G2}}`, etc.  
**Required:** These need to be mapped to specific trademark entries or removed/made static
