# 🔧 Bucket Creation Troubleshooting

## 🚨 **"Cannot find buckets" Error**

If you're seeing this error, here's what's happening and how to fix it:

### **Scenario 1: Planner has NO buckets at all**
**What you see:** "No buckets found in the selected planner"
**What it means:** The Microsoft Planner you selected is completely empty - no buckets exist yet.

**✅ Solutions:**
1. **Let the tool create them**: If your CSV has bucket names, the tool will offer to create buckets automatically
2. **Create manually first**: Go to Microsoft Planner and create at least one bucket
3. **Check permissions**: Make sure you can create buckets in this planner

### **Scenario 2: CSV bucket names don't match existing buckets**
**What you see:** Bucket lookup showing "❌ Not found" for your CSV bucket names
**What it means:** Your CSV has bucket names that don't exist in the planner, but the planner has other buckets.

**✅ Solutions:**
1. **Use bucket creation**: Check the "Enable bucket creation" box
2. **Select which buckets to create**: Use the checkboxes to choose which ones
3. **Click "Create Selected Buckets"**: The tool will create them for you

### **Scenario 3: Permission issues**
**What you see:** "Access denied" or "Failed to create buckets" errors
**What it means:** Your account doesn't have permission to create buckets.

**✅ Solutions:**
1. **Check planner membership**: Make sure you're a member of the planner
2. **Contact admin**: Ask your IT admin for bucket creation permissions
3. **Try a different planner**: Use a planner where you have full permissions

## 🔄 **Workflow After Bucket Creation**

After creating buckets, the tool will:
1. ✅ Show success message
2. 🔄 Automatically refresh the interface
3. 📋 Reload bucket mappings
4. 🚀 Continue to task creation

## 🎯 **Best Practices**

### **Before Uploading CSV:**
1. **Select your planner first** (this is now required!)
2. **Review available buckets** in the preview
3. **Ensure you have permissions** to create buckets if needed

### **CSV Format Tips:**
```csv
Title,Description,Start Date,Due Date,Assignee,Bucket Name,Status
"Task 1","Description","2024-01-10","2024-01-15","Shane Sweeney","Development","In Progress"
"Task 2","Description","2024-01-12","2024-01-18","Shane Sweeney","Marketing","Not Started"
```

### **Bucket Naming:**
- Use clear, descriptive names
- Be consistent with capitalization
- Avoid special characters
- Keep names under 50 characters

## 🐛 **Still Having Issues?**

### **Debug Steps:**
1. **Check the planner in Microsoft Planner web interface**:
   - Go to https://tasks.office.com
   - Open your planner
   - Verify you can create/edit buckets manually

2. **Verify CSV format**:
   - Ensure "Bucket Name" column exists
   - Check for typos in bucket names
   - Make sure cells aren't empty

3. **Test with simple data**:
   - Try with just 1-2 tasks
   - Use simple bucket names like "Test" or "Tasks"

4. **Clear browser cache**:
   - The tool uses session storage
   - Refresh the page to clear cached data

### **Common Fixes:**
- **Refresh the page** after creating buckets
- **Re-select the planner** if buckets don't appear
- **Try a different browser** if you have permission issues
- **Use the "Upload Different File" button** to restart the workflow

## 🆘 **Emergency Workarounds**

If bucket creation still isn't working:

1. **Manual creation**: Create buckets directly in Microsoft Planner first
2. **Simple approach**: Put all tasks in one bucket initially, organize later
3. **Remove bucket column**: Edit your CSV to remove the Bucket Name column temporarily
4. **Different planner**: Try with a planner where you have full admin rights

Remember: The bucket creation feature requires **write permissions** to the Microsoft Planner. If your organization restricts this, you may need to work with your IT administrator.