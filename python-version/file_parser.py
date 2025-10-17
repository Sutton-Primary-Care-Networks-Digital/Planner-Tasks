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
            st.write("3. Start Date (optional)")
            st.write("4. Due Date (optional)")
            st.write("5. Assignee (optional)")
            st.write("6. Bucket Name (optional)")
            st.write("7. Status (optional)")
        
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
        
        start_date_col = st.selectbox(
            "Select Start Date Column (optional):",
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
            help="Column containing names in format 'FirstName LastName (COMPANY)', 'FirstName LastName', or comma-separated multiple assignees"
        )
        
        bucket_col = st.selectbox(
            "Select Bucket Name Column (optional):",
            options=["None"] + available_columns,
            index=0,
            help="Column containing bucket names that will be matched against available buckets in the selected planner"
        )
        
        status_col = st.selectbox(
            "Select Status Column (optional):",
            options=["None"] + available_columns,
            index=0,
            help="Column containing status like 'In Progress', 'Complete', etc."
        )
        
        # Show assignee preview if column selected
        if assignee_col != "None":
            st.write("**Assignee Preview:**")
            assignee_sample = df[assignee_col].dropna().head(3).tolist()
            for i, name in enumerate(assignee_sample):
                st.write(f"- {name}")
            if len(assignee_sample) == 0:
                st.warning("No assignee data found in selected column")
        
        # Show bucket preview if column selected
        if bucket_col != "None":
            st.write("**Bucket Name Preview:**")
            bucket_sample = df[bucket_col].dropna().unique()[:5].tolist()
            for i, bucket in enumerate(bucket_sample):
                st.write(f"- {bucket}")
            if len(bucket_sample) == 0:
                st.warning("No bucket data found in selected column")
        
        # Show status preview if column selected
        if status_col != "None":
            st.write("**Status Preview:**")
            status_sample = df[status_col].dropna().unique()[:5].tolist()
            for i, status in enumerate(status_sample):
                st.write(f"- {status}")
            if len(status_sample) == 0:
                st.warning("No status data found in selected column")
        
        # Process the data
        if st.button("Process Data"):
            return self._process_mapped_data(
                df, title_col, description_col, start_date_col, due_date_col, assignee_col, bucket_col, status_col
            )
        
        return None
    
    def _process_mapped_data(self, df: pd.DataFrame, title_col: str, 
                           description_col: str, start_date_col: str, due_date_col: str, 
                           assignee_col: str, bucket_col: str, status_col: str) -> List[Dict[str, Any]]:
        """Process the mapped data into task objects"""
        tasks = []
        
        for index, row in df.iterrows():
            task = {
                "title": str(row[title_col]) if pd.notna(row[title_col]) else "",
                "description": "",
                "start_date": None,
                "due_date": None,
                "assignee": None,
                "bucket_name": None,
                "status": None
            }
            
            # Add description if column is selected and not "None"
            if description_col != "None" and pd.notna(row[description_col]):
                task["description"] = str(row[description_col])
            
            # Add start date if column is selected and not "None"
            if start_date_col != "None" and pd.notna(row[start_date_col]):
                normalized_start_date = self.normalize_date(row[start_date_col])
                task["start_date"] = normalized_start_date
            
            # Add due date if column is selected and not "None"
            if due_date_col != "None" and pd.notna(row[due_date_col]):
                # Use the new date normalization function
                normalized_date = self.normalize_date(row[due_date_col])
                task["due_date"] = normalized_date
            
            # Add assignee if column is selected and not "None"
            if assignee_col != "None" and pd.notna(row[assignee_col]):
                assignee_name = str(row[assignee_col]).strip()
                if assignee_name:
                    # Handle multiple assignees separated by commas
                    assignees = [name.strip() for name in assignee_name.split(',') if name.strip()]
                    task["assignee"] = assignee_name  # Keep original for display
                    task["assignees"] = assignees  # Store as list for processing
            
            # Add bucket name if column is selected and not "None"
            if bucket_col != "None" and pd.notna(row[bucket_col]):
                bucket_name = str(row[bucket_col]).strip()
                if bucket_name:
                    task["bucket_name"] = bucket_name
            
            # Add status if column is selected and not "None"
            if status_col != "None" and pd.notna(row[status_col]):
                status_name = str(row[status_col]).strip()
                if status_name:
                    task["status"] = status_name
            
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
                "Title": task["title"][:40] + "..." if len(task["title"]) > 40 else task["title"],
                "Description": task["description"][:25] + "..." if len(task["description"]) > 25 else task["description"],
                "Start Date": task["start_date"][:10] if task["start_date"] else "None",
                "Due Date": task["due_date"][:10] if task["due_date"] else "None",
                "Assignee": task["assignee"] if task["assignee"] else "None",
                "Bucket": task["bucket_name"] if task["bucket_name"] else "None",
                "Status": task["status"] if task["status"] else "None"
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
    
    def lookup_buckets(self, tasks: List[Dict[str, Any]], auth, access_token: str, plan_id: str) -> List[Dict[str, Any]]:
        """Lookup bucket names and add bucket information to tasks"""
        if not tasks or not plan_id:
            return tasks
        
        st.subheader("🗂️ Looking up Buckets")
        
        # Get buckets from the selected planner
        buckets = auth.get_planner_buckets(access_token, plan_id)
        if not buckets:
            st.warning("No buckets found in the selected planner")
            return tasks
        
        # Create bucket lookup dictionary (case-insensitive)
        bucket_lookup = {}
        for bucket in buckets:
            bucket_name = bucket['name']
            bucket_lookup[bucket_name.lower()] = {
                'id': bucket['id'],
                'name': bucket_name,
                'exact_match': bucket_name
            }
        
        # Get unique bucket names from tasks
        unique_bucket_names = set()
        for task in tasks:
            if task.get("bucket_name"):
                unique_bucket_names.add(task["bucket_name"])
        
        if not unique_bucket_names:
            st.info("No bucket names to lookup")
            return tasks
        
        # Show bucket mapping results
        st.write("**Bucket Mapping Results:**")
        bucket_mapping_results = {}
        
        for bucket_name in unique_bucket_names:
            # Try exact match first (case-insensitive)
            if bucket_name.lower() in bucket_lookup:
                matched_bucket = bucket_lookup[bucket_name.lower()]
                bucket_mapping_results[bucket_name] = matched_bucket
                st.success(f"✅ Found: {bucket_name} → {matched_bucket['name']}")
            else:
                # Try fuzzy matching
                best_match = None
                best_score = 0
                
                for available_bucket_name in bucket_lookup.keys():
                    # Simple substring matching
                    if bucket_name.lower() in available_bucket_name or available_bucket_name in bucket_name.lower():
                        score = min(len(bucket_name), len(available_bucket_name)) / max(len(bucket_name), len(available_bucket_name))
                        if score > best_score:
                            best_score = score
                            best_match = bucket_lookup[available_bucket_name]
                
                if best_match and best_score > 0.5:  # 50% similarity threshold
                    bucket_mapping_results[bucket_name] = best_match
                    st.warning(f"⚠️ Fuzzy match: {bucket_name} → {best_match['name']} (similarity: {best_score:.1%})")
                else:
                    bucket_mapping_results[bucket_name] = None
                    st.error(f"❌ Not found: {bucket_name}")
                    
        # Available buckets info
        with st.expander("📋 Available Buckets in Planner", expanded=False):
            st.write("Available buckets:")
            for bucket in buckets:
                st.write(f"- {bucket['name']}")
        
        # Handle bucket creation for missing buckets
        missing_buckets = [bucket_name for bucket_name, match in bucket_mapping_results.items() if not match]
        created_buckets = {}
        
        if missing_buckets:
            st.subheader("🔧 Create Missing Buckets")
            st.info(f"Found {len(missing_buckets)} bucket name(s) that don't exist in the planner.")
            
            # Global option to enable bucket creation
            create_buckets_enabled = st.checkbox(
                "✅ **Enable bucket creation**", 
                value=False,
                help="Enable this option to automatically create missing buckets before creating tasks",
                key="enable_bucket_creation"
            )
            
            if create_buckets_enabled:
                st.write("**Select which buckets to create:**")
                
                buckets_to_create = []
                
                # Create a more organized interface for bucket selection
                st.write("📝 **Bucket Creation Options:**")
                
                for i, bucket_name in enumerate(missing_buckets):
                    with st.container():
                        col1, col2, col3 = st.columns([1, 3, 2])
                        
                        with col1:
                            should_create = st.checkbox(
                                "✅ Create",
                                value=True,  # Default to checked
                                key=f"create_bucket_{bucket_name}",
                                help=f"Create bucket '{bucket_name}' in the planner"
                            )
                        
                        with col2:
                            st.write(f"🗂️ **{bucket_name}**")
                        
                        with col3:
                            task_count = sum(1 for task in tasks if task.get('bucket_name') == bucket_name)
                            st.info(f"{task_count} task(s)")
                        
                        if should_create:
                            buckets_to_create.append(bucket_name)
                        
                        if i < len(missing_buckets) - 1:  # Don't add separator after last item
                            st.write("---")
                
                # Convenience buttons
                col1, col2, col3 = st.columns(3)
                with col1:
                    if st.button("✅ Select All"):
                        # This will trigger a rerun with all checkboxes selected
                        for bucket_name in missing_buckets:
                            st.session_state[f"create_bucket_{bucket_name}"] = True
                        st.rerun()
                with col2:
                    if st.button("❌ Clear All"):
                        # This will trigger a rerun with all checkboxes unselected  
                        for bucket_name in missing_buckets:
                            st.session_state[f"create_bucket_{bucket_name}"] = False
                        st.rerun()
                
                # Create selected buckets
                if buckets_to_create:
                    st.markdown("---")
                    st.write(f"🎯 **Ready to create {len(buckets_to_create)} bucket(s):**")
                    for bucket_name in buckets_to_create:
                        task_count = sum(1 for task in tasks if task.get('bucket_name') == bucket_name)
                        st.write(f"- 🗂️ {bucket_name} ({task_count} task(s))")
                    
                    col1, col2 = st.columns([2, 1])
                    with col1:
                        if st.button("🔨 Create Selected Buckets", type="primary"):
                            with st.spinner("Creating buckets..."):
                                success_count = 0
                                for bucket_name in buckets_to_create:
                                    created_bucket = auth.create_bucket(access_token, plan_id, bucket_name)
                                    if created_bucket:
                                        success_count += 1
                                        created_buckets[bucket_name] = {
                                            'id': created_bucket['id'],
                                            'name': created_bucket['name'],
                                            'exact_match': created_bucket['name']
                                        }
                                        # Update the mapping results
                                        bucket_mapping_results[bucket_name] = created_buckets[bucket_name]
                            
                            # Show results and refresh
                            if success_count > 0:
                                st.success(f"✅ Successfully created {success_count} out of {len(buckets_to_create)} bucket(s)!")
                                st.info("🔄 Refreshing interface to show new buckets...")
                                # Clear bucket-related cache to force reload
                                cache_keys_to_clear = [
                                    key for key in st.session_state.keys() 
                                    if any(cache_term in key.lower() for cache_term in ['bucket_enriched', 'bucket_cache'])
                                ]
                                for key in cache_keys_to_clear:
                                    del st.session_state[key]
                                # Wait a moment for user to see the message
                                import time
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.error("❌ Failed to create any buckets. Please check permissions.")
                else:
                    st.info("📝 No buckets selected for creation.")
                
                # Show summary of what will happen
                if buckets_to_create:
                    with st.expander("🗒️ Bucket Creation Summary", expanded=False):
                        st.write("**Buckets that will be created:**")
                        for bucket_name in buckets_to_create:
                            task_count = sum(1 for task in tasks if task.get('bucket_name') == bucket_name)
                            st.write(f"- 🗂️ {bucket_name} ({task_count} task(s))")
                        
                        remaining_missing = [b for b in missing_buckets if b not in buckets_to_create]
                        if remaining_missing:
                            st.write("**Buckets that will use default bucket:**")
                            for bucket_name in remaining_missing:
                                task_count = sum(1 for task in tasks if task.get('bucket_name') == bucket_name)
                                st.write(f"- ❌ {bucket_name} ({task_count} task(s))")
        
        # Add bucket information to tasks
        enriched_tasks = []
        for task in tasks:
            enriched_task = task.copy()
            bucket_name = task.get("bucket_name")
            
            if bucket_name and bucket_name in bucket_mapping_results:
                matched_bucket = bucket_mapping_results[bucket_name]
                if matched_bucket:
                    enriched_task["bucket_info"] = {
                        "id": matched_bucket["id"],
                        "name": matched_bucket["name"],
                        "original_name": bucket_name,
                        "exact_match": matched_bucket["name"].lower() == bucket_name.lower()
                    }
                else:
                    enriched_task["bucket_lookup_failed"] = True
            
            enriched_tasks.append(enriched_task)
        
        # Show bucket lookup summary
        found_count = sum(1 for bucket_name, match in bucket_mapping_results.items() if match)
        total_count = len(unique_bucket_names)
        
        st.info(f"Bucket Lookup Complete: {found_count}/{total_count} found")
        
        return enriched_tasks
    
    def lookup_assignees(self, tasks: List[Dict[str, Any]], auth, access_token: str) -> List[Dict[str, Any]]:
        """Lookup assignees and add user information to tasks"""
        if not tasks:
            return tasks
        
        st.subheader("🔍 Looking up Assignees")
        
        # Get unique assignees (including individual names from multi-assignee tasks)
        unique_assignees = set()
        for task in tasks:
            if task.get("assignees"):
                # Handle multiple assignees
                for assignee in task["assignees"]:
                    unique_assignees.add(assignee)
            elif task.get("assignee"):
                # Handle single assignee (legacy)
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
            
            if task.get("assignees"):
                # Handle multiple assignees
                assignee_users = []
                failed_assignees = []
                
                for assignee_name in task["assignees"]:
                    if assignee_name in assignee_cache:
                        user = assignee_cache[assignee_name]
                        if user:
                            assignee_users.append({
                                "id": user["id"],
                                "displayName": user["displayName"],
                                "mail": user.get("mail", ""),
                                "originalName": assignee_name
                            })
                        else:
                            failed_assignees.append(assignee_name)
                
                if assignee_users:
                    enriched_task["assignee_users"] = assignee_users
                if failed_assignees:
                    enriched_task["assignee_lookup_failed_list"] = failed_assignees
                    
            elif task.get("assignee"):
                # Handle single assignee (legacy support)
                assignee_name = task["assignee"]
                if assignee_name in assignee_cache:
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