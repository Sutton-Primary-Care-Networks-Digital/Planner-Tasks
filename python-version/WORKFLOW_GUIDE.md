# Enhanced Microsoft Planner Task Creator - Workflow Guide

## 🔄 **New Improved Workflow**

### **Step 1: Upload File**
- Upload your CSV or Excel file
- Supported formats: `.csv`, `.xlsx`, `.xls`

### **Step 2: Select Planner FIRST** ⭐ **NEW**
- **Before processing the CSV**, you now select your target planner
- Shows available buckets in the selected planner for context
- This enables proper bucket mapping during CSV processing

### **Step 3: Column Mapping**
- Map your CSV columns to task fields:
  - **Title** (required)
  - **Description** (optional)
  - **Start Date** (optional) ⭐ **NEW**
  - **Due Date** (optional)
  - **Assignee** (optional) - supports multiple formats
  - **Bucket Name** (optional) ⭐ **NEW**
  - **Status** (optional) ⭐ **NEW**

### **Step 4: Bucket Lookup & Creation** ⭐ **NEW**
- Automatically matches CSV bucket names against the selected planner's buckets
- Shows exact matches and fuzzy matches with similarity scores
- **Bucket Creation Option**: For unmatched bucket names, offers to create them automatically
  - Global toggle to enable/disable bucket creation
  - Individual checkboxes for each missing bucket
  - Shows task count for each bucket
  - "Select All" and "Clear All" convenience buttons
- Falls back to default bucket for buckets not created

### **Step 5: Assignee Lookup**
- Looks up users in Microsoft Graph
- Supports multiple assignee formats:
  - "Shane Sweeney"
  - "Shane Sweeney (COMPANY)"
  - "Shane Sweeney, John Doe, Jane Smith" (comma-separated)

### **Step 6: Bucket Selection**
- Select default bucket for tasks without specific bucket assignments
- Shows summary of bucket mappings found
- Preview which CSV bucket names will match existing buckets

### **Step 7: Task Creation**
- Enhanced preview showing all fields including bucket assignments
- Creates tasks with:
  - Proper bucket assignments (individual or default)
  - Multiple assignees per task
  - Start and due dates
  - Status (mapped to completion percentage)

## 📋 **CSV Format Example**

```csv
Title,Description,Start Date,Due Date,Assignee,Bucket Name,Status
"Review documentation","Complete review of docs","2024-01-10","2024-01-15","Shane Sweeney","Documentation","In Progress"
"Multi-assignee task","Task for multiple people","2024-01-15","2024-01-20","Shane Sweeney, John Doe","Development","Not Started"
"Company format","Using company in name","2024-01-12","2024-01-16","Shane Sweeney (COMPANY)","Process","Complete"
```

## ⚡ **Key Benefits**

1. **Smart Bucket Mapping**: Automatically assigns tasks to the correct buckets based on CSV data
2. **Automatic Bucket Creation**: Creates missing buckets on-demand with user approval ⭐ **NEW**
3. **Flexible Assignee Support**: Handles any combination of assignee name formats
4. **Complete Date Support**: Both start dates and due dates with automatic parsing
5. **Status Integration**: Maps common status terms to Microsoft Planner progress values
6. **Early Planner Selection**: Enables bucket validation and mapping before processing

## 🔧 **Status Mapping**

- "Not Started" / "Not_Started" → 0%
- "In Progress" / "In_Progress" → 50%
- "Complete" / "Completed" → 100%

## 🗂️ **Bucket Assignment Logic**

1. **Exact Matching**: Tasks with bucket names that exactly match existing buckets are assigned to those buckets
2. **Fuzzy Matching**: Uses similarity scoring to match similar bucket names (e.g., "Dev" → "Development")
3. **Bucket Creation**: For unmatched bucket names, offers the option to create new buckets
4. **Default Bucket**: Tasks without bucket names or with unmatched/uncreated bucket names go to the selected default bucket

## 🔨 **Bucket Creation Feature** ⭐ **NEW**

When the tool finds bucket names in your CSV that don't exist in the selected planner:

### **Automatic Detection**
- Identifies all missing bucket names
- Shows how many tasks will use each missing bucket
- Provides a summary of what needs to be created

### **User Control**
- **Enable/Disable Toggle**: Master switch to enable bucket creation
- **Individual Selection**: Checkboxes for each missing bucket
- **Convenience Options**: "Select All" and "Clear All" buttons
- **Task Count Display**: Shows how many tasks will use each bucket

### **Creation Process**
- Creates buckets in real-time using Microsoft Graph API
- Shows creation progress and results
- Updates bucket mappings immediately after creation
- Refreshes the interface to show newly created buckets

### **Permissions**
- Requires write permissions to the Microsoft Planner
- Will show clear error messages if permissions are insufficient
- Falls back gracefully if bucket creation fails

## 🔄 **Navigation Options**

- **Upload Different File**: Clear all data and start over
- **Select Different Planner**: Keep file data but change planner selection
- **Back/Forward**: Navigate through the workflow steps

This enhanced workflow ensures optimal bucket mapping and provides a much more robust task creation experience!