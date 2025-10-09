"""
Microsoft Graph Authentication Module
Handles OAuth2 authentication with Microsoft Graph API using personal Microsoft account
"""

import msal
import streamlit as st
import requests
from typing import Optional, Dict, Any
import webbrowser
import time
import re

class GraphAuth:
    def __init__(self):
        # Try multiple client IDs that might work better with org restrictions
        self.client_ids = [
            "1950a258-227b-4e31-a9cf-717495945fc2",  # Microsoft Azure CLI client ID
            "04b07795-8ddb-461a-bbee-02f9e1bf7b46",  # Microsoft Graph Explorer client ID
            "1fec8e78-bce4-4aaf-ab1b-5451cc387264",  # Microsoft Graph PowerShell client ID
        ]
        self.tenant_id = "common"  # Use common tenant for personal accounts
        self.scopes = ["https://graph.microsoft.com/.default"]
    
    def authenticate_interactive(self) -> Optional[str]:
        """Authenticate using interactive browser flow with multiple client IDs"""
        authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        
        # Try each client ID until one works
        for i, client_id in enumerate(self.client_ids):
            try:
                st.info(f"**Trying authentication method {i+1} of {len(self.client_ids)}...**")
                
                app = msal.PublicClientApplication(
                    client_id,
                    authority=authority
                )
                
                # Try to get token silently first
                accounts = app.get_accounts()
                if accounts:
                    result = app.acquire_token_silent(self.scopes, account=accounts[0])
                    if result and "access_token" in result:
                        return result["access_token"]
                
                # Use interactive browser flow
                st.info("**Opening browser for authentication...**")
                
                result = app.acquire_token_interactive(
                    scopes=self.scopes,
                    prompt="select_account"
                )
                
                if result and "access_token" in result:
                    st.success(f"✅ Authentication successful with method {i+1}!")
                    return result["access_token"]
                else:
                    error_message = result.get("error_description", "Unknown error") if result else "No result returned"
                    st.warning(f"❌ Method {i+1} failed: {error_message}")
                    
                    # Check for specific admin restriction error
                    if result and "error" in result:
                        error_code = result.get("error")
                        if "53003" in str(error_code) or "admin" in error_message.lower():
                            st.error("🚫 **Admin restrictions detected!** This organization blocks this application.")
                            st.warning("**Try these solutions:**")
                            st.write("1. Use a **personal Microsoft account** (@outlook.com, @hotmail.com, @live.com)")
                            st.write("2. Contact your IT admin about Microsoft Graph access")
                            st.write("3. Try a different browser or incognito mode")
                            
            except Exception as e:
                st.warning(f"❌ Method {i+1} failed with exception: {str(e)}")
                continue
        
        st.error("❌ All authentication methods failed")
        return None
    
    def get_planners(self, access_token: str) -> Optional[list]:
        """Get list of planners from Microsoft Graph"""
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            # Get teams first
            response = requests.get(
                "https://graph.microsoft.com/v1.0/me/joinedTeams",
                headers=headers
            )
            
            if response.status_code == 401:
                st.error("❌ Authentication expired. Please sign in again.")
                if "access_token" in st.session_state:
                    del st.session_state.access_token
                st.rerun()
                return None
            elif response.status_code == 403:
                st.error("❌ Access denied. You may not have permission to access teams.")
                # Try fallback to groups
                return self._get_planners_fallback(access_token)
            elif response.status_code == 200:
                teams = response.json().get("value", [])
                planners = []
                
                for team in teams:
                    group_id = team.get("id")
                    group_name = team.get("displayName", "Unknown Group")
                    
                    # Get plans for this group
                    plans_response = requests.get(
                        f"https://graph.microsoft.com/v1.0/groups/{group_id}/planner/plans",
                        headers=headers
                    )
                    
                    if plans_response.status_code == 200:
                        plans = plans_response.json().get("value", [])
                        for plan in plans:
                            # Filter out planners with "NHS.net" in the title
                            if "NHS.net" not in plan["title"]:
                                planners.append({
                                    "id": plan["id"],
                                    "title": plan["title"],
                                    "groupId": group_id,
                                    "groupName": group_name
                                })
                
                return planners
            else:
                st.error(f"Failed to get planners: {response.text}")
                return None
                
        except Exception as e:
            st.error(f"Error getting planners: {str(e)}")
            return None
    
    def _get_planners_fallback(self, access_token: str) -> Optional[list]:
        """Fallback method to get planners via groups"""
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            # Try to get groups
            response = requests.get(
                "https://graph.microsoft.com/v1.0/me/memberOf?$filter=groupTypes/any(c:c eq 'Unified')",
                headers=headers
            )
            
            if response.status_code == 200:
                groups = response.json().get("value", [])
                planners = []
                
                for group in groups:
                    group_id = group.get("id")
                    group_name = group.get("displayName", "Unknown Group")
                    
                    # Get plans for this group
                    plans_response = requests.get(
                        f"https://graph.microsoft.com/v1.0/groups/{group_id}/planner/plans",
                        headers=headers
                    )
                    
                    if plans_response.status_code == 200:
                        plans = plans_response.json().get("value", [])
                        for plan in plans:
                            # Filter out planners with "NHS.net" in the title
                            if "NHS.net" not in plan["title"]:
                                planners.append({
                                    "id": plan["id"],
                                    "title": plan["title"],
                                    "groupId": group_id,
                                    "groupName": group_name
                                })
                
                return planners
            else:
                st.error("Could not access Microsoft Planner. You may need proper permissions.")
                return None
                
        except Exception as e:
            st.error(f"Error in fallback method: {str(e)}")
            return None
    
    def get_planner_buckets(self, access_token: str, plan_id: str) -> Optional[list]:
        """Get buckets for a specific planner"""
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            response = requests.get(
                f"https://graph.microsoft.com/v1.0/planner/plans/{plan_id}/buckets",
                headers=headers
            )
            
            if response.status_code == 401:
                st.error("❌ Authentication expired. Please sign in again.")
                if "access_token" in st.session_state:
                    del st.session_state.access_token
                st.rerun()
                return None
            elif response.status_code == 200:
                return response.json().get("value", [])
            else:
                st.error(f"Failed to get buckets: {response.text}")
                return None
                
        except Exception as e:
            st.error(f"Error getting buckets: {str(e)}")
            return None
    
    def parse_display_name(self, display_name: str) -> Optional[Dict[str, str]]:
        """Parse display name in format 'FirstName LastName (COMPANY)' and extract components"""
        if not display_name:
            return None
            
        # Pattern to match "FirstName LastName (COMPANY)" format
        pattern = r'^(.+?)\s+(.+?)\s*\((.+?)\)\s*$'
        match = re.match(pattern, display_name.strip())
        
        if match:
            first_name = match.group(1).strip()
            last_name = match.group(2).strip()
            company = match.group(3).strip()
            
            return {
                "firstName": first_name,
                "lastName": last_name,
                "company": company,
                "fullName": f"{first_name} {last_name}",
                "displayName": display_name
            }
        
        # If pattern doesn't match, try simpler patterns
        # Try "FirstName LastName" without company
        simple_pattern = r'^(.+?)\s+(.+?)$'
        simple_match = re.match(simple_pattern, display_name.strip())
        
        if simple_match:
            first_name = simple_match.group(1).strip()
            last_name = simple_match.group(2).strip()
            
            return {
                "firstName": first_name,
                "lastName": last_name,
                "company": "",
                "fullName": f"{first_name} {last_name}",
                "displayName": display_name
            }
        
        # If no pattern matches, return the original as full name
        return {
            "firstName": "",
            "lastName": "",
            "company": "",
            "fullName": display_name,
            "displayName": display_name
        }
    
    def search_user(self, access_token: str, assignee_name: str) -> Optional[Dict[str, Any]]:
        """Search for a user by display name using Microsoft Graph"""
        if not assignee_name or assignee_name.strip() == "":
            return None
            
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        # Parse the display name to extract components
        parsed_name = self.parse_display_name(assignee_name)
        if not parsed_name:
            return None
        
        try:
            # Search strategies in order of preference
            search_terms = [
                parsed_name["fullName"],  # "FirstName LastName"
                parsed_name["displayName"],  # Original full string
                parsed_name["firstName"],  # Just first name
                parsed_name["lastName"]   # Just last name
            ]
            
            for search_term in search_terms:
                if not search_term:
                    continue
                    
                # Try different search approaches
                search_queries = [
                    f"https://graph.microsoft.com/v1.0/users?$filter=displayName eq '{search_term}'&$select=id,displayName,mail,userPrincipalName",
                    f"https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'{search_term}')&$select=id,displayName,mail,userPrincipalName",
                    f"https://graph.microsoft.com/v1.0/users?$search=\"displayName:{search_term}\"&$select=id,displayName,mail,userPrincipalName"
                ]
                
                for query_url in search_queries:
                    try:
                        response = requests.get(query_url, headers=headers)
                        
                        if response.status_code == 200:
                            users = response.json().get("value", [])
                            
                            # Look for exact or close matches
                            for user in users:
                                user_display_name = user.get("displayName", "")
                                
                                # Check for exact match first
                                if user_display_name.lower() == parsed_name["fullName"].lower():
                                    return user
                                
                                # Check if user display name contains the search components
                                if (parsed_name["firstName"].lower() in user_display_name.lower() and 
                                    parsed_name["lastName"].lower() in user_display_name.lower()):
                                    return user
                            
                            # If exact matches not found, return first result for exact search
                            if users and "eq" in query_url:
                                return users[0]
                                
                    except Exception as e:
                        continue  # Try next search approach
            
            return None
            
        except Exception as e:
            st.warning(f"Error searching for user '{assignee_name}': {str(e)}")
            return None
    
    def assign_task(self, access_token: str, task_id: str, user_id: str) -> bool:
        """Assign a task to a user"""
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            # First, get the task to retrieve its ETag
            task_response = requests.get(
                f"https://graph.microsoft.com/v1.0/planner/tasks/{task_id}",
                headers=headers
            )
            
            if task_response.status_code != 200:
                st.warning(f"Could not get task for assignment: {task_response.text}")
                return False
            
            # Get the etag from the response headers
            etag = task_response.headers.get('ETag', '')
            if not etag:
                st.warning("Could not get ETag for task assignment")
                return False
            
            # Add If-Match header with the etag
            headers["If-Match"] = etag
            
            # Create assignment data - FIXED: orderHint should be a string, not object
            assignments = {user_id: {"@odata.type": "microsoft.graph.plannerAssignment", "orderHint": " !"}}
            
            # Update the task with assignment
            assignment_data = {"assignments": assignments}
            
            response = requests.patch(
                f"https://graph.microsoft.com/v1.0/planner/tasks/{task_id}",
                headers=headers,
                json=assignment_data
            )
            
            # Better error handling - 204 is success for PATCH operations
            if response.status_code in [200, 204]:
                return True
            else:
                st.warning(f"Assignment failed: {response.status_code} - {response.text}")
                return False
            
        except Exception as e:
            st.warning(f"Error assigning task to user: {str(e)}")
            return False

    def create_task(self, access_token: str, plan_id: str, bucket_id: str, 
                   title: str, description: str = "", due_date: str = None, 
                   assignee: str = None) -> Optional[Dict[str, Any]]:
        """Create a new task in Microsoft Planner with optional assignee"""
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        # Step 1: Create the basic task
        task_data = {
            "planId": plan_id,
            "bucketId": bucket_id,
            "title": title
        }
        
        if due_date:
            task_data["dueDateTime"] = due_date
        
        try:
            # Create the task
            response = requests.post(
                "https://graph.microsoft.com/v1.0/planner/tasks",
                headers=headers,
                json=task_data
            )
            
            if response.status_code == 401:
                st.error("❌ Authentication expired. Please sign in again.")
                if "access_token" in st.session_state:
                    del st.session_state.access_token
                st.rerun()
                return None
            elif response.status_code == 403:
                st.error("❌ Access denied. You may not have permission to create tasks in this planner.")
                st.warning("**Possible solutions:**")
                st.write("1. Make sure you have write permissions in the selected planner")
                st.write("2. Check if you're a member of the group that contains this planner")
                st.write("3. Try selecting a different planner or bucket")
                return None
            elif response.status_code != 201:
                st.error(f"Failed to create task: {response.text}")
                return None
            
            # Get the created task
            task = response.json()
            task_id = task["id"]
            
            # Step 2: Update the task with description if provided
            if description:
                self._update_task_description(access_token, task_id, description)
            
            # Step 3: Assign the task if assignee is provided
            if assignee:
                user = self.search_user(access_token, assignee)
                if user:
                    assignment_success = self.assign_task(access_token, task_id, user["id"])
                    if assignment_success:
                        # Add assignee info to the returned task
                        task["assignedUser"] = {
                            "id": user["id"],
                            "displayName": user["displayName"],
                            "mail": user.get("mail", "")
                        }
                    # Note: We don't show error messages here as they're often misleading
                # Note: We don't show error messages here as they're often misleading
            
            return task
                
        except Exception as e:
            st.error(f"Error creating task: {str(e)}")
            return None
    
    def _update_task_description(self, access_token: str, task_id: str, description: str) -> bool:
        """Update task description using the proper /details endpoint"""
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            # Step 1: Get the task details to get the ETag
            get_response = requests.get(
                f"https://graph.microsoft.com/v1.0/planner/tasks/{task_id}/details",
                headers=headers
            )
            
            if get_response.status_code != 200:
                return False
            
            # Get the etag from the response headers
            etag = get_response.headers.get('ETag', '')
            if not etag:
                return False
            
            # Add If-Match header with the etag
            headers["If-Match"] = etag
            
            # Step 2: Update the task description using the /details endpoint
            update_data = {
                "description": description
            }
            
            response = requests.patch(
                f"https://graph.microsoft.com/v1.0/planner/tasks/{task_id}/details",
                headers=headers,
                json=update_data
            )
            
            return response.status_code in [200, 204]
                
        except Exception as e:
            return False