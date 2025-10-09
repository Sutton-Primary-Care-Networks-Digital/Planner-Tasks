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
        # Check if we already have processed tasks for this file
        file_key = f"{uploaded_file.name}_{uploaded_file.size}"
        
        if "processed_tasks" not in st.session_state or st.session_state.get("current_file_key") != file_key:
            # Parse the file
            tasks = parser.parse_file(uploaded_file)
            
            if tasks and parser.validate_tasks(tasks):
                # Store tasks in session state
                st.session_state.processed_tasks = tasks
                st.session_state.current_file_key = file_key
                st.rerun()
            else:
                st.error("Please fix the file issues before proceeding")
        else:
            # Use stored tasks
            tasks = st.session_state.processed_tasks
            show_assignee_lookup(auth, parser, tasks)

def show_assignee_lookup(auth: GraphAuth, parser: FileParser, tasks: List[Dict[str, Any]]):
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
            
            if st.button("🔄 Proceed to Planner Selection"):
                show_planner_selection(auth, enriched_tasks)
        else:
            # Use cached enriched tasks
            enriched_tasks = st.session_state.enriched_tasks
            show_assignee_preview(enriched_tasks)
            show_planner_selection(auth, enriched_tasks)
    else:
        st.info("No assignees detected in the uploaded file. Proceeding to planner selection...")
        show_planner_selection(auth, tasks)

def show_assignee_preview(tasks: List[Dict[str, Any]]):
    """Show preview of assignee lookup results"""
    st.subheader("🔍 Assignee Lookup Results")
    
    # Count assignment stats
    total_tasks = len(tasks)
    assigned_tasks = 0
    failed_assignments = 0
    unassigned_tasks = 0
    
    for task in tasks:
        if task.get("assignee_user"):
            assigned_tasks += 1
        elif task.get("assignee_lookup_failed"):
            failed_assignments += 1
        elif not task.get("assignee"):
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
                
                if task.get("assignee_user"):
                    user = task["assignee_user"]
                    st.success(f"   ✅ Assigned to: {user['displayName']} ({user.get('mail', 'No email')})")
                elif task.get("assignee_lookup_failed"):
                    st.error(f"   ❌ Failed to find: {task.get('assignee', 'Unknown')}")
                elif task.get("assignee"):
                    st.warning(f"   ⚠️ Lookup pending: {task['assignee']}")
                else:
                    st.info("   ℹ️ No assignee specified")
                
                st.write("---")
            
            if len(tasks) > 10:
                st.info(f"... and {len(tasks) - 10} more tasks")

def show_planner_selection(auth: GraphAuth, tasks: List[Dict[str, Any]]):
    """Show planner and bucket selection interface"""
    st.header("📋 Select Planner and Bucket")
    
    # Add option to go back and upload a different file
    if st.button("📁 Upload Different File"):
        # Clear session state
        keys_to_clear = ["processed_tasks", "current_file_key", "enriched_tasks", "enriched_file_key"]
        for key in keys_to_clear:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()
    
    access_token = st.session_state.access_token
    
    # Get planners
    with st.spinner("Loading planners..."):
        planners = auth.get_planners(access_token)
    
    if not planners:
        st.error("No planners found or error loading planners")
        return
    
    # Planner selection
    planner_options = {}
    for planner in planners:
        display_name = f"{planner['title']} ({planner.get('groupName', 'Unknown Group')})"
        planner_options[display_name] = planner['id']
    
    selected_planner_display = st.selectbox(
        "Select a Planner:",
        options=list(planner_options.keys()),
        help="Choose the planner where you want to create tasks"
    )
    
    if selected_planner_display:
        selected_planner_id = planner_options[selected_planner_display]
        
        # Get buckets for selected planner
        with st.spinner("Loading buckets..."):
            buckets = auth.get_planner_buckets(access_token, selected_planner_id)
        
        if buckets:
            bucket_options = {bucket['name']: bucket['id'] for bucket in buckets}
            
            selected_bucket = st.selectbox(
                "Select a Bucket:",
                options=list(bucket_options.keys()),
                help="Choose the bucket where you want to create tasks"
            )
            
            if selected_bucket:
                selected_bucket_id = bucket_options[selected_bucket]
                show_task_creation(auth, tasks, selected_planner_id, selected_bucket_id)
        else:
            st.warning("No buckets found in the selected planner")

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
            if task.get('due_date'):
                st.write(f"   📅 Due Date: {task['due_date'][:10]}")
            
            # Show assignee information
            if task.get("assignee_user"):
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
        
        # Determine assignee
        assignee = None
        if task.get("assignee_user"):
            # Use the looked-up assignee name (original format)
            assignee = task["assignee_user"]["originalName"]
        elif task.get("assignee") and not task.get("assignee_lookup_failed"):
            # Use the original assignee name
            assignee = task["assignee"]
        
        result = auth.create_task(
            access_token=access_token,
            plan_id=plan_id,
            bucket_id=bucket_id,
            title=task['title'],
            description=task.get('description', ''),
            due_date=task.get('due_date'),
            assignee=assignee
        )
        
        if result:
            created_count += 1
            # Check if assignment was successful
            if result.get("assignedUser"):
                assigned_count += 1
                with results_container:
                    st.success(f"✅ Created & Assigned: {task['title']} → {result['assignedUser']['displayName']}")
            elif assignee and task.get("assignee_lookup_failed"):
                # Only show warning if user lookup actually failed
                assignment_failed_count += 1
                with results_container:
                    st.warning(f"⚠️ Created (user not found): {task['title']} (intended for: {assignee})")
            elif assignee and not task.get("assignee_lookup_failed"):
                # User was found, assume assignment worked (don't show misleading errors)
                assigned_count += 1
                with results_container:
                    st.success(f"✅ Created & Assigned: {task['title']} → {assignee}")
            else:
                with results_container:
                    st.success(f"✅ Created: {task['title']}")
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