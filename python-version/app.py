"""
Microsoft Planner Task Creator
Main Streamlit application for creating tasks from CSV/Excel files
"""

import streamlit as st
from graph_auth import GraphAuth
from file_parser import FileParser
from typing import List, Dict, Any, Optional

# Page configuration
st.set_page_config(
    page_title="Microsoft Planner Task Creator",
    page_icon="📋",
    layout="wide"
)

def main():
    st.title("📋 Microsoft Planner Task Creator")
    st.markdown("Create tasks in Microsoft Planner from CSV or Excel files with assignee support")
    
    # Initialize components
    auth = GraphAuth()
    parser = FileParser()
    
    # Check if user is authenticated
    if "access_token" not in st.session_state:
        show_authentication(auth)
    else:
        show_main_interface(auth, parser)

def show_authentication(auth: GraphAuth):
    """Show authentication interface"""
    st.header("🔐 Authentication Required")
    
    st.info("""
    **Simple Setup - No Azure Configuration Required!**
    
    This tool uses your Microsoft account to access Microsoft Planner.
    """)
    
    # Show troubleshooting info
    with st.expander("⚠️ Having Authentication Issues?", expanded=False):
        st.write("""
        **If you're getting blocked due to admin restrictions:**
        
        1. **Try a personal Microsoft account** (@outlook.com, @hotmail.com, @live.com)
        2. **If using a work account**, contact your IT admin about Microsoft Graph access
        3. **Make sure you have access** to Microsoft Planner in your organization
        4. **Try using a different browser** or incognito mode
        
        **Alternative Solutions:**
        - Use a personal Microsoft account that has access to Microsoft Planner
        - Ask your IT admin to whitelist Microsoft Graph Explorer applications
        - Use a different device/network that doesn't have the same restrictions
        """)
    
    # Authentication button
    if st.button("🔑 Sign in with Microsoft Account", type="primary"):
        try:
            with st.spinner("Starting authentication..."):
                access_token = auth.authenticate_interactive()
                if access_token:
                    st.session_state.access_token = access_token
                    st.success("✅ Authentication successful!")
                    st.rerun()
                else:
                    st.error("❌ Authentication failed. Please try the troubleshooting tips above.")
        except Exception as e:
            st.error(f"❌ Authentication failed: {str(e)}")
            st.warning("**This is likely due to admin restrictions.** Try using a personal Microsoft account instead of your work account.")

def show_main_interface(auth: GraphAuth, parser: FileParser):
    """Show main application interface"""
    # Add sign out button
    col1, col2 = st.columns([3, 1])
    with col1:
        st.header("📁 Upload File")
    with col2:
        if st.button("🚪 Sign Out", help="Sign out and re-authenticate"):
            # Clear session state
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose a CSV or Excel file",
        type=['csv', 'xlsx', 'xls'],
        help="Upload a file with columns for title, description, due date, and optional assignee"
    )
    
    if uploaded_file is not None:
        # First, select planner before processing CSV
        selected_planner_info = show_planner_selection_first(auth)
        
        if selected_planner_info:
            # Check if we already have processed tasks for this file with this planner
            file_key = f"{uploaded_file.name}_{uploaded_file.size}_{selected_planner_info['id']}"
            
            if "processed_tasks" not in st.session_state or st.session_state.get("current_file_key") != file_key:
                # Parse the file with planner context
                tasks = parser.parse_file(uploaded_file)
                
                if tasks and parser.validate_tasks(tasks):
                    # Store tasks and planner info in session state
                    st.session_state.processed_tasks = tasks
                    st.session_state.current_file_key = file_key
                    st.session_state.selected_planner_info = selected_planner_info
                    st.rerun()
                else:
                    st.error("Please fix the file issues before proceeding")
            else:
                # Use stored tasks and planner info
                tasks = st.session_state.processed_tasks
                show_file_processing_workflow(auth, parser, tasks, selected_planner_info)

def show_planner_selection_first(auth: GraphAuth) -> Optional[Dict[str, Any]]:
    """Show planner selection before CSV processing"""
    st.header("📋 Select Planner")
    st.info("Please select your Microsoft Planner first. This helps with bucket mapping when processing your CSV file.")
    
    access_token = st.session_state.access_token
    
    # Get planners
    with st.spinner("Loading planners..."):
        planners = auth.get_planners(access_token)
    
    if not planners:
        st.error("No planners found or error loading planners")
        return None
    
    # Planner selection
    planner_options = {}
    for planner in planners:
        display_name = f"{planner['title']} ({planner.get('groupName', 'Unknown Group')})"
        planner_options[display_name] = {
            'id': planner['id'],
            'title': planner['title'],
            'groupName': planner.get('groupName', 'Unknown Group'),
            'display_name': display_name
        }
    
    selected_planner_display = st.selectbox(
        "Select a Planner:",
        options=list(planner_options.keys()),
        help="Choose the planner where you want to create tasks. This selection helps with bucket mapping.",
        key="planner_selection_first"
    )
    
    if selected_planner_display:
        selected_planner_info = planner_options[selected_planner_display]
        
        # Show bucket preview for context
        with st.spinner("Loading buckets for preview..."):
            buckets = auth.get_planner_buckets(access_token, selected_planner_info['id'])
        
        if buckets:
            with st.expander(f"📂 Available Buckets in '{selected_planner_info['title']}':", expanded=False):
                st.write("Your CSV bucket names will be matched against these buckets:")
                for bucket in buckets:
                    st.write(f"- {bucket['name']}")
        
        return selected_planner_info
    
    return None

def show_file_processing_workflow(auth: GraphAuth, parser: FileParser, tasks: List[Dict[str, Any]], planner_info: Dict[str, Any]):
    """Show the file processing workflow with planner context"""
    st.header("📁 Processing Workflow")
    
    # Show selected planner info
    st.success(f"📋 Selected Planner: **{planner_info['display_name']}**")
    
    # Check if tasks have bucket names for lookup
    has_bucket_names = any(task.get("bucket_name") for task in tasks)
    
    if has_bucket_names:
        st.info("🗂️ Bucket names detected in your CSV. These will be matched against the selected planner's buckets.")
        
        # Perform bucket lookup immediately with planner context
        access_token = st.session_state.access_token
        bucket_cache_key = f"bucket_{planner_info['id']}_{st.session_state.get('current_file_key')}"
        
        if "bucket_enriched_tasks" not in st.session_state or st.session_state.get("bucket_cache_key") != bucket_cache_key:
            # Perform bucket lookup
            bucket_enriched_tasks = parser.lookup_buckets(tasks, auth, access_token, planner_info['id'])
            
            # Store bucket enriched tasks
            st.session_state.bucket_enriched_tasks = bucket_enriched_tasks
            st.session_state.bucket_cache_key = bucket_cache_key
            tasks = bucket_enriched_tasks
        else:
            # Use cached bucket enriched tasks
            tasks = st.session_state.bucket_enriched_tasks
    
    # Now proceed with assignee lookup
    show_assignee_lookup(auth, parser, tasks, planner_info)

def show_assignee_lookup(auth: GraphAuth, parser: FileParser, tasks: List[Dict[str, Any]], planner_info: Dict[str, Any]):
    """Show assignee lookup interface"""
    st.header("👥 Assignee Lookup")
    
    # Check if tasks have assignees
    has_assignees = any(task.get("assignee") for task in tasks)
    
    if has_assignees:
        st.info("Tasks with assignees detected. Looking up users in Microsoft Graph...")
        
        # Check if we already have enriched tasks
        if "enriched_tasks" not in st.session_state or st.session_state.get("current_file_key") != st.session_state.get("enriched_file_key"):
            # Perform assignee lookup
            access_token = st.session_state.access_token
            enriched_tasks = parser.lookup_assignees(tasks, auth, access_token)
            
            # Store enriched tasks
            st.session_state.enriched_tasks = enriched_tasks
            st.session_state.enriched_file_key = st.session_state.get("current_file_key")
            
            # Show lookup results
            show_assignee_preview(enriched_tasks)
            
            if st.button("🔄 Proceed to Bucket Selection"):
                show_bucket_selection(auth, enriched_tasks, planner_info)
        else:
            # Use cached enriched tasks
            enriched_tasks = st.session_state.enriched_tasks
            show_assignee_preview(enriched_tasks)
            show_bucket_selection(auth, enriched_tasks, planner_info)
    else:
        st.info("No assignees detected in the uploaded file. Proceeding to bucket selection...")
        show_bucket_selection(auth, tasks, planner_info)

def show_assignee_preview(tasks: List[Dict[str, Any]]):
    """Show preview of assignee lookup results"""
    st.subheader("🔍 Assignee Lookup Results")
    
    # Count assignment stats
    total_tasks = len(tasks)
    assigned_tasks = 0
    failed_assignments = 0
    unassigned_tasks = 0
    
    for task in tasks:
        if task.get("assignee_users"):  # Multiple assignees
            assigned_tasks += len(task["assignee_users"])
            if task.get("assignee_lookup_failed_list"):
                failed_assignments += len(task["assignee_lookup_failed_list"])
        elif task.get("assignee_user"):  # Single assignee (legacy)
            assigned_tasks += 1
        elif task.get("assignee_lookup_failed"):
            failed_assignments += 1
        elif not task.get("assignee") and not task.get("assignees"):
            unassigned_tasks += 1
    
    # Show statistics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Tasks", total_tasks)
    with col2:
        st.metric("Successfully Assigned", assigned_tasks)
    with col3:
        st.metric("Assignment Failed", failed_assignments)
    with col4:
        st.metric("Unassigned", unassigned_tasks)
    
    # Show detailed preview
    if assigned_tasks > 0 or failed_assignments > 0:
        with st.expander("📋 Detailed Assignment Preview", expanded=True):
            for i, task in enumerate(tasks[:10]):  # Show first 10 tasks
                st.write(f"**{i+1}. {task['title']}**")
                
                if task.get("assignee_users"):  # Multiple assignees
                    users = task["assignee_users"]
                    assignee_names = [f"{user['displayName']} ({user.get('mail', 'No email')})" for user in users]
                    st.success(f"   ✅ Assigned to: {', '.join(assignee_names)}")
                    if task.get("assignee_lookup_failed_list"):
                        failed = task["assignee_lookup_failed_list"]
                        st.error(f"   ❌ Failed to find: {', '.join(failed)}")
                elif task.get("assignee_user"):  # Single assignee (legacy)
                    user = task["assignee_user"]
                    st.success(f"   ✅ Assigned to: {user['displayName']} ({user.get('mail', 'No email')})")
                elif task.get("assignee_lookup_failed"):
                    st.error(f"   ❌ Failed to find: {task.get('assignee', 'Unknown')}")
                elif task.get("assignee") or task.get("assignees"):
                    assignee_text = task.get("assignee") or ', '.join(task.get("assignees", []))
                    st.warning(f"   ⚠️ Lookup pending: {assignee_text}")
                else:
                    st.info("   ℹ️ No assignee specified")
                
                st.write("---")
            
            if len(tasks) > 10:
                st.info(f"... and {len(tasks) - 10} more tasks")

def show_bucket_selection(auth: GraphAuth, tasks: List[Dict[str, Any]], planner_info: Dict[str, Any]):
    """Show bucket selection interface (planner already selected)"""
    st.header("🗂️ Select Default Bucket")
    
    # Show selected planner
    st.info(f"Creating tasks in: **{planner_info['display_name']}**")
    
    # Add option to go back and select different planner/file
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📁 Upload Different File"):
            # Clear session state
            keys_to_clear = ["processed_tasks", "current_file_key", "enriched_tasks", "enriched_file_key", "bucket_enriched_tasks", "bucket_cache_key", "selected_planner_info"]
            for key in keys_to_clear:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()
    with col2:
        if st.button("📋 Select Different Planner"):
            # Clear session state except file data
            keys_to_clear = ["enriched_tasks", "enriched_file_key", "bucket_enriched_tasks", "bucket_cache_key", "selected_planner_info"]
            for key in keys_to_clear:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()
    
    access_token = st.session_state.access_token
    
    # Get buckets for the already selected planner
    with st.spinner("Loading buckets..."):
        buckets = auth.get_planner_buckets(access_token, planner_info['id'])
    
    if not buckets:
        st.warning("🗂️ No buckets found in the selected planner")
        st.info("💡 **Don't worry!** We can create buckets for you.")
        
        # Check if tasks have bucket names that we can create
        unique_bucket_names = set(task.get("bucket_name") for task in tasks if task.get("bucket_name"))
        
        if unique_bucket_names:
            st.subheader("🔧 Create Buckets from CSV")
            st.success(f"Found {len(unique_bucket_names)} unique bucket names in your CSV!")
            
            # Show bucket creation interface
            create_buckets_enabled = st.checkbox(
                "✅ **Create buckets from CSV**", 
                value=True,
                help="Create buckets based on the bucket names found in your CSV file",
                key="enable_bucket_creation_empty_planner"
            )
            
            if create_buckets_enabled:
                st.write("**Buckets that will be created:**")
                
                buckets_to_create = []
                for bucket_name in sorted(unique_bucket_names):
                    task_count = sum(1 for task in tasks if task.get('bucket_name') == bucket_name)
                    
                    col1, col2, col3 = st.columns([1, 3, 2])
                    with col1:
                        should_create = st.checkbox(
                            "✅",
                            value=True,
                            key=f"create_bucket_empty_{bucket_name}",
                            help=f"Create bucket '{bucket_name}'"
                        )
                    with col2:
                        st.write(f"🗂️ **{bucket_name}**")
                    with col3:
                        st.info(f"{task_count} task(s)")
                    
                    if should_create:
                        buckets_to_create.append(bucket_name)
                
                # Create buckets
                if buckets_to_create:
                    st.write(f"🎯 Ready to create {len(buckets_to_create)} bucket(s)")
                    
                    if st.button("🔨 Create All Buckets", type="primary"):
                        with st.spinner("Creating buckets..."):
                            created_count = 0
                            for bucket_name in buckets_to_create:
                                created_bucket = auth.create_bucket(access_token, planner_info['id'], bucket_name)
                                if created_bucket:
                                    created_count += 1
                            
                            if created_count > 0:
                                st.success(f"✅ Successfully created {created_count} bucket(s)!")
                                st.info("🔄 Refreshing interface to show new buckets...")
                                st.write("🎉 **Great!** Your buckets are now ready for task creation.")
                                # Clear bucket cache to force reload
                                cache_keys_to_clear = [
                                    key for key in st.session_state.keys() 
                                    if any(cache_term in key.lower() for cache_term in ['bucket_enriched', 'bucket_cache'])
                                ]
                                for key in cache_keys_to_clear:
                                    if key in st.session_state:
                                        del st.session_state[key]
                                # Wait a moment for user to see the message
                                import time
                                time.sleep(1.5)
                                st.rerun()
                            else:
                                st.error("❌ Failed to create buckets. Please check permissions.")
        else:
            # No bucket names in CSV
            st.info("📝 Your CSV doesn't contain bucket names.")
            st.write("**Options:**")
            st.write("1. Add a 'Bucket Name' column to your CSV and re-upload")
            st.write("2. Create a bucket manually in Microsoft Planner first")
            st.write("3. Contact your admin to create initial buckets")
            
            with st.expander("🔧 Need Help?", expanded=False):
                st.markdown("""
                **Troubleshooting Steps:**
                1. Check if you can create buckets manually in Microsoft Planner
                2. Verify you're a member of the selected planner  
                3. Try using a different planner where you have admin rights
                4. Contact your IT admin for Microsoft Planner permissions
                
                **Quick Test:** Try creating a bucket manually at https://tasks.office.com
                """)
        
        return
    
    bucket_options = {bucket['name']: bucket['id'] for bucket in buckets}
    
    # Check if tasks have individual bucket assignments
    has_bucket_names = any(task.get("bucket_name") for task in tasks)
    
    if has_bucket_names:
        st.success("✨ Great! Your CSV contains bucket names. Tasks will be created in their specified buckets when found.")
        st.write("**Default bucket** (for tasks without matching bucket names):")
    else:
        st.write("**Select the bucket** where all tasks will be created:")
    
    selected_bucket = st.selectbox(
        "Default Bucket:",
        options=list(bucket_options.keys()),
        help="Tasks will be created in this bucket, unless they have specific bucket names that match other buckets"
    )
    
    if selected_bucket:
        selected_bucket_id = bucket_options[selected_bucket]
        
        # Show bucket assignment summary
        if has_bucket_names:
            unique_bucket_names = set(task.get("bucket_name") for task in tasks if task.get("bucket_name"))
            
            st.write("**Bucket Assignment Summary:**")
            st.write(f"- 🗂️ Default bucket: **{selected_bucket}**")
            st.write(f"- 🔍 CSV bucket names found: {len(unique_bucket_names)}")
            
            # Show which CSV bucket names exist
            if unique_bucket_names:
                with st.expander("CSV Bucket Names", expanded=False):
                    for bucket_name in sorted(unique_bucket_names):
                        # Check if bucket exists in the planner buckets
                        bucket_exists = bucket_name.lower() in [b['name'].lower() for b in buckets]
                        # Check if task has bucket_info (meaning it was found or created)
                        task_with_bucket = next((task for task in tasks if task.get('bucket_name') == bucket_name), None)
                        has_bucket_info = task_with_bucket and task_with_bucket.get('bucket_info')
                        
                        if bucket_exists:
                            st.write(f"✅ {bucket_name} (existing bucket)")
                        elif has_bucket_info:
                            st.write(f"🆕 {bucket_name} (newly created)")
                        else:
                            st.write(f"❌ {bucket_name} (will use default)")
        
        show_task_creation(auth, tasks, planner_info['id'], selected_bucket_id)

def show_task_creation(auth: GraphAuth, tasks: List[Dict[str, Any]], 
                      plan_id: str, bucket_id: str):
    """Show task creation interface with assignee preview"""
    st.header("🚀 Create Tasks")
    
    st.write(f"**Ready to create {len(tasks)} tasks**")
    
    # Show enhanced task preview with assignees
    with st.expander("📋 Preview Tasks with Assignments", expanded=False):
        for i, task in enumerate(tasks[:10]):  # Show first 10 tasks
            st.write(f"**{i+1}.** {task['title']}")
            if task.get('description'):
                st.write(f"   📝 Description: {task['description'][:100]}{'...' if len(task.get('description', '')) > 100 else ''}")
            if task.get('start_date'):
                st.write(f"   🚀 Start Date: {task['start_date'][:10]}")
            if task.get('due_date'):
                st.write(f"   📅 Due Date: {task['due_date'][:10]}")
            if task.get('status'):
                st.write(f"   📊 Status: {task['status']}")
            
            # Show bucket information
            if task.get('bucket_info'):
                bucket = task['bucket_info']
                if bucket['exact_match']:
                    st.write(f"   🗂️ **Bucket:** {bucket['name']}")
                else:
                    st.write(f"   🗂️ **Bucket:** {bucket['original_name']} → {bucket['name']} (mapped)")
            elif task.get('bucket_lookup_failed'):
                st.write(f"   ❌ **Bucket not found:** {task.get('bucket_name', 'Unknown')}")
            elif task.get('bucket_name'):
                st.write(f"   🗂️ **Bucket:** {task['bucket_name']} (pending lookup)")
            
            # Show assignee information
            if task.get("assignee_users"):  # Multiple assignees
                users = task["assignee_users"]
                assignee_names = [user['displayName'] for user in users]
                st.write(f"   👥 **Assigned to:** {', '.join(assignee_names)}")
                if task.get("assignee_lookup_failed_list"):
                    failed = task["assignee_lookup_failed_list"]
                    st.write(f"   ❌ **Assignment Failed:** {', '.join(failed)}")
            elif task.get("assignee_user"):  # Single assignee (legacy)
                user = task["assignee_user"]
                st.write(f"   👤 **Assigned to:** {user['displayName']}")
            elif task.get("assignee_lookup_failed"):
                st.write(f"   ❌ **Assignment Failed:** {task.get('assignee', 'Unknown')}")
            elif task.get("assignee"):
                st.write(f"   ⚠️ **Assignee:** {task['assignee']} (lookup pending)")
            else:
                st.write(f"   ℹ️ **No assignee**")
            
            st.write("---")
        
        if len(tasks) > 10:
            st.info(f"... and {len(tasks) - 10} more tasks")
    
    # Show confirmation modal
    show_confirmation_modal(auth, tasks, plan_id, bucket_id)

def show_confirmation_modal(auth: GraphAuth, tasks: List[Dict[str, Any]], 
                           plan_id: str, bucket_id: str):
    """Show confirmation modal with assignee information"""
    
    # Get planner and bucket names for display
    access_token = st.session_state.access_token
    
    # Get planner name
    planners = auth.get_planners(access_token)
    planner_name = "Unknown Planner"
    if planners:
        for planner in planners:
            if planner['id'] == plan_id:
                planner_name = f"{planner['title']} ({planner.get('groupName', 'Unknown Group')})"
                break
    
    # Get bucket name
    buckets = auth.get_planner_buckets(access_token, plan_id)
    bucket_name = "Unknown Bucket"
    if buckets:
        for bucket in buckets:
            if bucket['id'] == bucket_id:
                bucket_name = bucket['name']
                break
    
    # Count assignment statistics
    total_tasks = len(tasks)
    assigned_tasks = sum(1 for task in tasks if task.get("assignee_user"))
    failed_assignments = sum(1 for task in tasks if task.get("assignee_lookup_failed"))
    unassigned_tasks = total_tasks - assigned_tasks - failed_assignments
    
    # Confirmation modal
    with st.container():
        st.markdown("---")
        st.subheader("⚠️ Confirmation Required")
        
        # Display target information
        col1, col2 = st.columns(2)
        with col1:
            st.write("**📋 Planner:**")
            st.write(planner_name)
        with col2:
            st.write("**🗂️ Bucket:**")
            st.write(bucket_name)
        
        # Display task and assignment statistics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Tasks", total_tasks)
        with col2:
            st.metric("With Assignees", assigned_tasks)
        with col3:
            st.metric("Failed Assignments", failed_assignments)
        with col4:
            st.metric("Unassigned", unassigned_tasks)
        
        # Warning message
        st.warning(f"""
        **Please verify the above information is correct before proceeding.**
        
        This action will create {total_tasks} tasks in the selected planner and bucket.
        {f"{assigned_tasks} tasks will be assigned to users." if assigned_tasks > 0 else ""}
        {f"{failed_assignments} tasks have failed assignee lookups and will be created without assignments." if failed_assignments > 0 else ""}
        """)
        
        # Confirmation checkbox
        st.markdown("---")
        
        # Checkbox for confirmation
        confirmation_text = f"I confirm that I want to create {total_tasks} tasks in '{planner_name}' → '{bucket_name}'"
        confirmed = st.checkbox(confirmation_text, key="task_confirmation")
        
        # Create tasks button (enabled only when checkbox is checked)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if confirmed:
                if st.button("🚀 Create All Tasks", type="primary", key="create_tasks_btn"):
                    create_tasks_with_progress(auth, tasks, plan_id, bucket_id)
            else:
                st.button("🚀 Create All Tasks", disabled=True, key="disabled_create_tasks_btn")
                st.caption("Please check the confirmation box above to enable this button")

def create_tasks_with_progress(auth: GraphAuth, tasks: List[Dict[str, Any]], 
                              plan_id: str, bucket_id: str):
    """Create tasks with progress tracking and assignee support"""
    access_token = st.session_state.access_token
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    results_container = st.container()
    
    created_count = 0
    failed_count = 0
    assigned_count = 0
    assignment_failed_count = 0
    failed_tasks = []
    
    for i, task in enumerate(tasks):
        status_text.text(f"Creating task {i+1} of {len(tasks)}: {task['title']}")
        
        # Determine assignees
        assignees = []
        if task.get("assignee_users"):  # Multiple assignees
            assignees = [user["originalName"] for user in task["assignee_users"]]
        elif task.get("assignee_user"):  # Single assignee (legacy)
            assignees = [task["assignee_user"]["originalName"]]
        elif task.get("assignees"):  # Original assignee list
            assignees = task["assignees"]
        elif task.get("assignee") and not task.get("assignee_lookup_failed"):
            assignees = [task["assignee"]]
        
        # Determine bucket ID (use task-specific bucket if available)
        task_bucket_id = bucket_id  # Default bucket
        if task.get("bucket_info"):
            task_bucket_id = task["bucket_info"]["id"]
        
        result = auth.create_task(
            access_token=access_token,
            plan_id=plan_id,
            bucket_id=task_bucket_id,
            title=task['title'],
            description=task.get('description', ''),
            due_date=task.get('due_date'),
            start_date=task.get('start_date'),
            assignees=assignees if assignees else None,
            status=task.get('status')
        )
        
        if result:
            created_count += 1
            
            # Prepare display message with bucket info
            bucket_info = ""
            if task.get("bucket_info"):
                bucket_name = task["bucket_info"]["name"]
                if not task["bucket_info"]["exact_match"]:
                    bucket_info = f" in {bucket_name}"
                else:
                    bucket_info = f" in {bucket_name}"
            
            # Check if assignment was successful
            if result.get("assignedUsers"):  # Multiple assignees
                assigned_count += len(result["assignedUsers"])
                assignee_names = [user['displayName'] for user in result["assignedUsers"]]
                with results_container:
                    st.success(f"✅ Created & Assigned: {task['title']}{bucket_info} → {', '.join(assignee_names)}")
            elif result.get("assignedUser"):  # Single assignee (legacy)
                assigned_count += 1
                with results_container:
                    st.success(f"✅ Created & Assigned: {task['title']}{bucket_info} → {result['assignedUser']['displayName']}")
            elif assignees and (task.get("assignee_lookup_failed") or task.get("assignee_lookup_failed_list")):
                # Only show warning if user lookup actually failed
                assignment_failed_count += 1
                failed_names = task.get("assignee_lookup_failed_list", [task.get("assignee", "Unknown")])
                with results_container:
                    st.warning(f"⚠️ Created (user not found): {task['title']}{bucket_info} (intended for: {', '.join(failed_names)})")
            elif assignees and not (task.get("assignee_lookup_failed") or task.get("assignee_lookup_failed_list")):
                # Users were found, assume assignment worked
                assigned_count += len(assignees)
                with results_container:
                    st.success(f"✅ Created & Assigned: {task['title']}{bucket_info} → {', '.join(assignees)}")
            else:
                with results_container:
                    st.success(f"✅ Created: {task['title']}{bucket_info}")
        else:
            failed_count += 1
            failed_tasks.append(task)
            with results_container:
                st.error(f"❌ Failed: {task['title']}")
        
        # Update progress
        progress_bar.progress((i + 1) / len(tasks))
    
    # Final results
    status_text.text("Task creation completed!")
    
    st.header("📊 Results Summary")
    
    # Display comprehensive statistics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("✅ Created", created_count)
    with col2:
        st.metric("👤 Assigned", assigned_count)
    with col3:
        st.metric("⚠️ Assignment Failed", assignment_failed_count)
    with col4:
        st.metric("❌ Failed", failed_count)
    
    if created_count > 0:
        st.success(f"Successfully created {created_count} out of {len(tasks)} tasks!")
        if assigned_count > 0:
            st.success(f"Successfully assigned {assigned_count} tasks to users!")
        if assignment_failed_count > 0:
            st.warning(f"{assignment_failed_count} tasks were created but could not be assigned.")
    
    if failed_count > 0:
        st.error(f"Failed to create {failed_count} tasks")
        with st.expander("Failed Tasks Details"):
            for task in failed_tasks:
                st.write(f"- {task['title']}")
    
    # Option to create another batch
    if st.button("🔄 Create More Tasks"):
        # Clear relevant session state
        keys_to_clear = ["processed_tasks", "current_file_key", "enriched_tasks", "enriched_file_key"]
        for key in keys_to_clear:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()

if __name__ == "__main__":
    main()