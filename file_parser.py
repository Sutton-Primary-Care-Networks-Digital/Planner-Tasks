"""
File Parser Module
Handles CSV and Excel file parsing for task data
"""

import pandas as pd
import streamlit as st
from typing import List, Dict, Any, Optional
import io
from dateutil import parser
from datetime import datetime

class FileParser:
    def __init__(self):
        self.required_columns = ["title", "description", "due_date", "assignee"]
        self.optional_columns = ["title", "description", "due_date", "assignee"]
    
    def normalize_date(self, date_str: str) -> Optional[str]:
        """Convert various date formats to ISO 8601 format required by Microsoft Planner"""
        if not date_str or pd.isna(date_str) or str(date_str).strip() == "":
            return None
        
        try:
            # Parse the date string using dateutil (handles many formats)
            parsed_date = parser.parse(str(date_str))
            
            # Convert to ISO 8601 format with UTC timezone
            if parsed_date.tzinfo is None:
                # If no timezone info, assume UTC
                iso_date = parsed_date.replace(tzinfo=None).isoformat() + "Z"
            else:
                # Convert to UTC and format
                iso_date = parsed_date.astimezone().isoformat()
                if not iso_date.endswith('Z'):
                    iso_date = iso_date.replace('+00:00', 'Z')
            
            return iso_date
            
        except (ValueError, TypeError) as e:
            st.warning(f"Could not parse date '{date_str}': {str(e)}")
            return None
    
    def parse_file(self, uploaded_file) -> Optional[List[Dict[str, Any]]]:
        """Parse uploaded CSV or Excel file and return list of tasks"""
        try:
            # Determine file type and read accordingly
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            elif uploaded_file.name.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(uploaded_file)
            else:
                st.error("Unsupported file format. Please upload a CSV or Excel file.")
                return None
            
            # Display file info
            st.info(f"File loaded: {uploaded_file.name} with {len(df)} rows")
            
            # Show column mapping interface
            return self._map_columns(df)
            
        except Exception as e:
            st.error(f"Error parsing file: {str(e)}")
            return None
    
    def _map_columns(self, df: pd.DataFrame) -> Optional[List[Dict[str, Any]]]:
        """Map file columns to required task fields"""
        st.subheader("Column Mapping")
        st.write("Map your file columns to the required task fields:")
        
        # Get available columns
        available_columns = list(df.columns)
        
        # Create column mapping interface
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Available Columns:**")
            for i, col in enumerate(available_columns):
                st.write(f"{i+1}. {col}")
        
        with col2:
            st.write("**Required Fields:**")
            st.write("1. Title (required)")
            st.write("2. Description (optional)")
            st.write("3. Due Date (optional)")
            st.write("4. Assignee (optional)")
        
        # Column mapping
        title_col = st.selectbox(
            "Select Title Column:",
            options=available_columns,
            index=0
        )
        
        description_col = st.selectbox(
            "Select Description Column (optional):",
            options=["None"] + available_columns,
            index=0
        )
        
        due_date_col = st.selectbox(
            "Select Due Date Column (optional):",
            options=["None"] + available_columns,
            index=0
        )
        
        assignee_col = st.selectbox(
            "Select Assignee Column (optional):",
            options=["None"] + available_columns,
            index=0,
            help="Column containing names in format 'FirstName LastName (COMPANY)' or 'FirstName LastName'"
        )
        
        # Show assignee preview if column selected
        if assignee_col != "None":
            st.write("**Assignee Preview:**")
            assignee_sample = df[assignee_col].dropna().head(3).tolist()
            for i, name in enumerate(assignee_sample):
                st.write(f"- {name}")
            if len(assignee_sample) == 0:
                st.warning("No assignee data found in selected column")
        
        # Process the data
        if st.button("Process Data"):
            return self._process_mapped_data(
                df, title_col, description_col, due_date_col, assignee_col
            )
        
        return None
    
    def _process_mapped_data(self, df: pd.DataFrame, title_col: str, 
                           description_col: str, due_date_col: str, assignee_col: str) -> List[Dict[str, Any]]:
        """Process the mapped data into task objects"""
        tasks = []
        
        for index, row in df.iterrows():
            task = {
                "title": str(row[title_col]) if pd.notna(row[title_col]) else "",
                "description": "",
                "due_date": None,
                "assignee": None
            }
            
            # Add description if column is selected and not "None"
            if description_col != "None" and pd.notna(row[description_col]):
                task["description"] = str(row[description_col])
            
            # Add due date if column is selected and not "None"
            if due_date_col != "None" and pd.notna(row[due_date_col]):
                # Use the new date normalization function
                normalized_date = self.normalize_date(row[due_date_col])
                task["due_date"] = normalized_date
            
            # Add assignee if column is selected and not "None"
            if assignee_col != "None" and pd.notna(row[assignee_col]):
                assignee_name = str(row[assignee_col]).strip()
                if assignee_name:
                    task["assignee"] = assignee_name
            
            # Only add tasks with non-empty titles
            if task["title"].strip():
                tasks.append(task)
        
        st.success(f"Processed {len(tasks)} tasks from {len(df)} rows")
        
        # Show preview of processed tasks
        if tasks:
            self._show_task_preview(tasks)
        
        return tasks
    
    def _show_task_preview(self, tasks: List[Dict[str, Any]]):
        """Show preview of processed tasks with assignee information"""
        st.subheader("Task Preview")
        
        # Create preview dataframe
        preview_data = []
        for task in tasks[:5]:  # Show first 5 tasks
            preview_task = {
                "Title": task["title"][:50] + "..." if len(task["title"]) > 50 else task["title"],
                "Description": task["description"][:30] + "..." if len(task["description"]) > 30 else task["description"],
                "Due Date": task["due_date"][:10] if task["due_date"] else "None",
                "Assignee": task["assignee"] if task["assignee"] else "None"
            }
            preview_data.append(preview_task)
        
        preview_df = pd.DataFrame(preview_data)
        st.dataframe(preview_df, use_container_width=True)
        
        if len(tasks) > 5:
            st.info(f"Showing first 5 of {len(tasks)} tasks")
        
        # Show assignee statistics
        assignee_stats = self._get_assignee_statistics(tasks)
        if assignee_stats:
            st.write("**Assignee Statistics:**")
            for assignee, count in assignee_stats.items():
                st.write(f"- {assignee}: {count} task(s)")
    
    def _get_assignee_statistics(self, tasks: List[Dict[str, Any]]) -> Dict[str, int]:
        """Get statistics about assignees in the tasks"""
        assignee_counts = {}
        unassigned_count = 0
        
        for task in tasks:
            assignee = task.get("assignee")
            if assignee:
                assignee_counts[assignee] = assignee_counts.get(assignee, 0) + 1
            else:
                unassigned_count += 1
        
        if unassigned_count > 0:
            assignee_counts["[Unassigned]"] = unassigned_count
        
        return assignee_counts
    
    def lookup_assignees(self, tasks: List[Dict[str, Any]], auth, access_token: str) -> List[Dict[str, Any]]:
        """Lookup assignees and add user information to tasks"""
        if not tasks:
            return tasks
        
        st.subheader("🔍 Looking up Assignees")
        
        # Get unique assignees
        unique_assignees = set()
        for task in tasks:
            if task.get("assignee"):
                unique_assignees.add(task["assignee"])
        
        if not unique_assignees:
            st.info("No assignees to lookup")
            return tasks
        
        # Lookup each unique assignee
        assignee_cache = {}
        progress_bar = st.progress(0)
        
        for i, assignee_name in enumerate(unique_assignees):
            progress_bar.progress((i + 1) / len(unique_assignees))
            
            if assignee_name not in assignee_cache:
                user = auth.search_user(access_token, assignee_name)
                assignee_cache[assignee_name] = user
                
                if user:
                    st.success(f"✅ Found: {assignee_name} → {user['displayName']}")
                else:
                    st.warning(f"❌ Not found: {assignee_name}")
        
        progress_bar.empty()
        
        # Add user information to tasks
        enriched_tasks = []
        for task in tasks:
            enriched_task = task.copy()
            assignee_name = task.get("assignee")
            
            if assignee_name and assignee_name in assignee_cache:
                user = assignee_cache[assignee_name]
                if user:
                    enriched_task["assignee_user"] = {
                        "id": user["id"],
                        "displayName": user["displayName"],
                        "mail": user.get("mail", ""),
                        "originalName": assignee_name
                    }
                else:
                    enriched_task["assignee_lookup_failed"] = True
            
            enriched_tasks.append(enriched_task)
        
        # Show lookup summary
        found_count = sum(1 for assignee, user in assignee_cache.items() if user)
        total_count = len(unique_assignees)
        
        st.info(f"Assignee Lookup Complete: {found_count}/{total_count} found")
        
        return enriched_tasks
    
    def validate_tasks(self, tasks: List[Dict[str, Any]]) -> bool:
        """Validate that all tasks have required fields"""
        if not tasks:
            st.error("No tasks to validate")
            return False
        
        valid_tasks = []
        invalid_tasks = []
        
        for i, task in enumerate(tasks):
            if not task.get("title", "").strip():
                invalid_tasks.append(f"Row {i+1}: Missing title")
            else:
                valid_tasks.append(task)
        
        if invalid_tasks:
            st.warning(f"Found {len(invalid_tasks)} invalid tasks:")
            for invalid in invalid_tasks:
                st.write(f"- {invalid}")
        
        st.info(f"Valid tasks: {len(valid_tasks)}")
        
        return len(valid_tasks) == len(tasks)