import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import anthropic
import os
from datetime import datetime, timedelta
import io
import openpyxl
from openpyxl import Workbook
import json
import uuid
import re
import time

# Set page configuration
st.set_page_config(
    page_title="Due Diligence Tracker",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state variables if they don't exist
if 'checklist_data' not in st.session_state:
    st.session_state.checklist_data = None
if 'project_name' not in st.session_state:
    st.session_state.project_name = None
if 'project_config' not in st.session_state:
    st.session_state.project_config = None
if 'filtered_data' not in st.session_state:
    st.session_state.filtered_data = None
if 'claude_insights' not in st.session_state:
    st.session_state.claude_insights = {}
if 'last_analyzed' not in st.session_state:
    st.session_state.last_analyzed = None

# Function to initialize Anthropic client
def init_claude_client():
    api_key = st.session_state.get('anthropic_api_key', '')
    if api_key:
        return anthropic.Anthropic(api_key=api_key)
    return None

# Functions for Excel processing
def parse_excel_checklist(uploaded_file):
    """Parse Excel file and extract checklist data"""
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)
        
        # Check if the DataFrame is empty
        if df.empty:
            st.error("The uploaded file does not contain any data.")
            return None
        
        # Clean up column names - strip whitespace and convert to lowercase
        df.columns = [str(col).strip().lower() for col in df.columns]
        
        # Try to identify key columns based on common patterns
        item_col = next((col for col in df.columns if any(keyword in col for keyword in ['item', 'task', 'activity', 'description', 'check'])), df.columns[0])
        status_col = next((col for col in df.columns if any(keyword in col for keyword in ['status', 'state', 'complete', 'progress'])), None)
        category_col = next((col for col in df.columns if any(keyword in col for keyword in ['category', 'group', 'type', 'area'])), None)
        owner_col = next((col for col in df.columns if any(keyword in col for keyword in ['owner', 'responsible', 'assignee', 'assigned'])), None)
        due_date_col = next((col for col in df.columns if any(keyword in col for keyword in ['due', 'date', 'deadline', 'target'])), None)
        
        # Create a standardized DataFrame
        checklist_data = pd.DataFrame()
        
        # Add item description column (required)
        checklist_data['item_description'] = df[item_col].astype(str)
        
        # Add status column (create if not exists)
        if status_col in df.columns:
            checklist_data['status'] = df[status_col]
        else:
            checklist_data['status'] = "Not Started"
            
        # Add category column (create if not exists)
        if category_col in df.columns:
            checklist_data['category'] = df[category_col]
        else:
            # Try to infer categories from item descriptions
            checklist_data['category'] = "General"
            
        # Add owner column (create if not exists)
        if owner_col in df.columns:
            checklist_data['owner'] = df[owner_col]
        else:
            checklist_data['owner'] = "Unassigned"
            
        # Add due date column (create if not exists) with robust date handling
        if due_date_col in df.columns:
            # Handle date conversion safely
            due_dates = []
            for date_val in df[due_date_col]:
                if pd.isna(date_val):
                    # Use a default date for missing values
                    due_dates.append(pd.Timestamp.now() + pd.Timedelta(days=14))
                else:
                    try:
                        # Try to convert to datetime, with various approaches
                        if isinstance(date_val, (int, float)):
                            # Handle Excel's numeric date format
                            due_dates.append(pd.to_datetime('1899-12-30') + pd.Timedelta(days=int(date_val)))
                        else:
                            # Try normal parsing
                            due_dates.append(pd.to_datetime(date_val, errors='coerce'))
                    except:
                        # If all else fails, use a default date
                        due_dates.append(pd.Timestamp.now() + pd.Timedelta(days=14))
            
            checklist_data['due_date'] = due_dates
        else:
            # Set default due date to 2 weeks from now
            checklist_data['due_date'] = pd.Timestamp.now() + pd.Timedelta(days=14)
            
        # Add priority column (create if not exists)
        checklist_data['priority'] = "Medium"
        
        # Add notes column
        checklist_data['notes'] = ""
        
        # Add unique ID for each item
        checklist_data['id'] = [str(uuid.uuid4()) for _ in range(len(checklist_data))]
        
        # Add last updated timestamp
        checklist_data['last_updated'] = pd.Timestamp.now()
        
        # Convert all text columns to string type
        for col in checklist_data.columns:
            if checklist_data[col].dtype == 'object' and col not in ['due_date', 'last_updated']:
                checklist_data[col] = checklist_data[col].astype(str)
                
        # Handle NaN values
        checklist_data = checklist_data.fillna({
            'item_description': 'No description provided',
            'status': 'Not Started',
            'category': 'General',
            'owner': 'Unassigned',
            'priority': 'Medium',
            'notes': ''
        })
        
        return checklist_data
        
    except Exception as e:
        st.error(f"Error parsing Excel file: {str(e)}")
        return None

def export_to_excel(df, filename="due_diligence_checklist_export.xlsx"):
    """Export DataFrame to Excel file for download"""
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer

# Functions for Claude AI integration
def analyze_checklist_with_claude(client, checklist_data, project_name):
    """Use Claude to analyze the checklist and provide insights"""
    if client is None:
        st.warning("Please enter your Claude API key in the settings to enable AI insights.")
        return {}
    
    # Prepare the checklist data for analysis
    checklist_text = checklist_data[['item_description', 'category', 'status', 'priority', 'owner']].to_string(index=False)
    
    # Count items by status
    status_counts = checklist_data['status'].value_counts().to_dict()
    
    # Count items by category
    category_counts = checklist_data['category'].value_counts().to_dict()
    
    # Count items by priority
    priority_counts = checklist_data['priority'].value_counts().to_dict()
    
    # Calculate percentage complete
    total_items = len(checklist_data)
    completed_items = sum(1 for status in checklist_data['status'] if 'complete' in status.lower())
    completion_percentage = (completed_items / total_items) * 100 if total_items > 0 else 0
    
    # Build the prompt for Claude
    prompt = f"""
    You are an expert project management consultant analyzing a due diligence checklist for project: {project_name}.
    
    Here is a summary of the checklist data:
    - Total items: {total_items}
    - Completion percentage: {completion_percentage:.2f}%
    - Status breakdown: {status_counts}
    - Category breakdown: {category_counts}
    - Priority breakdown: {priority_counts}
    
    Here's the detailed checklist:
    {checklist_text}
    
    Please provide the following insights:
    1. OVERALL_ASSESSMENT: A brief assessment of the overall project status based on the checklist.
    2. KEY_RISKS: Identify 3-5 potential risks or critical items that need attention.
    3. SUGGESTIONS: Provide 3-5 actionable suggestions to improve progress.
    4. CATEGORIES_ANALYSIS: Brief analysis of each category's status.
    5. NEXT_STEPS: Recommended next steps to advance the due diligence process.
    
    Format your response as JSON with the above keys.
    """
    
    try:
        # Call Claude API
        message = client.messages.create(
            model="claude-3-haiku-20240307",
            max_tokens=1500,
            temperature=0.2,
            system="You are an expert project management consultant specializing in due diligence processes. Provide concise, actionable insights based on checklist data.",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        
        # Extract JSON from response
        response_text = message.content[0].text
        
        # Find JSON in the response (it might be wrapped in code blocks)
        json_match = re.search(r'```json\n(.*?)\n```', response_text, re.DOTALL)
        if json_match:
            json_str = json_match.group(1)
        else:
            # If not in code block, try to extract JSON directly
            json_str = response_text
        
        # Parse the JSON
        insights = json.loads(json_str)
        
        return insights
        
    except Exception as e:
        st.error(f"Error getting insights from Claude: {str(e)}")
        # Return default structure with error message
        return {
            "OVERALL_ASSESSMENT": f"Error analyzing checklist: {str(e)}",
            "KEY_RISKS": ["Unable to analyze risks due to API error"],
            "SUGGESTIONS": ["Check your API key and try again"],
            "CATEGORIES_ANALYSIS": "Analysis unavailable",
            "NEXT_STEPS": ["Retry analysis when API is available"]
        }

def get_claude_recommendations(client, checklist_item):
    """Get Claude's recommendations for a specific checklist item"""
    if client is None:
        return "Claude API key not configured. Please add your API key in settings."
    
    try:
        prompt = f"""
        You are an expert project management consultant. 
        Please provide specific, actionable recommendations for the following due diligence checklist item:
        
        Item: {checklist_item['item_description']}
        Category: {checklist_item['category']}
        Current Status: {checklist_item['status']}
        Priority: {checklist_item['priority']}
        Owner: {checklist_item['owner']}
        
        Provide 2-3 specific actions that would help complete this item effectively.
        Keep your response under 200 words and focus on practical next steps.
        """
        
        message = client.messages.create(
            model="claude-3-haiku-20240307",
            max_tokens=300,
            temperature=0.2,
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        
        return message.content[0].text
        
    except Exception as e:
        return f"Error getting recommendations: {str(e)}"

# Main Streamlit app layout
def main():
    # Sidebar with settings and project selection
    with st.sidebar:
        st.image("https://via.placeholder.com/150x60?text=DueDiligence", width=150)
        st.title("Project Settings")
        
        # Claude API settings
        st.subheader("Claude AI Integration")
        api_key = st.text_input("Claude API Key", value=st.session_state.get('anthropic_api_key', ''), type="password")
        if api_key:
            st.session_state.anthropic_api_key = api_key
            client = init_claude_client()
            if client:
                st.success("Claude API connected", icon="✅")
        else:
            st.info("Enter Claude API key to enable AI insights")
            client = None
        
        st.divider()
        
        # Project settings
        st.subheader("Project Configuration")
        
        # Option to create new project or use existing
        new_project = st.text_input("Create new project", placeholder="Enter project name")
        
        if new_project:
            st.session_state.project_name = new_project
            st.session_state.checklist_data = None
            st.success(f"Created new project: {new_project}")
        
        # Upload Excel checklist
        st.subheader("Upload Checklist")
        uploaded_file = st.file_uploader("Upload due diligence checklist", type=["xlsx", "xls"])
        
        if uploaded_file is not None:
            # Parse the uploaded Excel file
            checklist_data = parse_excel_checklist(uploaded_file)
            
            if checklist_data is not None:
                st.session_state.checklist_data = checklist_data
                st.session_state.filtered_data = checklist_data.copy()
                st.success(f"Uploaded checklist with {len(checklist_data)} items")
                
                # If no project name is set, use the filename
                if not st.session_state.project_name:
                    st.session_state.project_name = uploaded_file.name.split(".")[0]
        
        st.divider()
        
        # Project navigation and filters
        if st.session_state.checklist_data is not None:
            st.subheader("Filters")
            
            # Status filter
            status_options = ["All"] + list(st.session_state.checklist_data['status'].unique())
            selected_status = st.selectbox("Status", status_options)
            
            # Category filter
            category_options = ["All"] + list(st.session_state.checklist_data['category'].unique())
            selected_category = st.selectbox("Category", category_options)
            
            # Priority filter
            priority_options = ["All"] + list(st.session_state.checklist_data['priority'].unique())
            selected_priority = st.selectbox("Priority", priority_options)
            
            # Owner filter
            owner_options = ["All"] + list(st.session_state.checklist_data['owner'].unique())
            selected_owner = st.selectbox("Owner", owner_options)
            
            # Apply filters
            filtered_data = st.session_state.checklist_data.copy()
            
            if selected_status != "All":
                filtered_data = filtered_data[filtered_data['status'] == selected_status]
            
            if selected_category != "All":
                filtered_data = filtered_data[filtered_data['category'] == selected_category]
                
            if selected_priority != "All":
                filtered_data = filtered_data[filtered_data['priority'] == selected_priority]
                
            if selected_owner != "All":
                filtered_data = filtered_data[filtered_data['owner'] == selected_owner]
                
            st.session_state.filtered_data = filtered_data
            
            st.markdown(f"**Showing {len(filtered_data)} of {len(st.session_state.checklist_data)} items**")

    # Main content area
    if st.session_state.project_name:
        st.title(f"Due Diligence Tracker: {st.session_state.project_name}")
        
        # Tabs for different views
        tab1, tab2, tab3, tab4 = st.tabs(["Dashboard", "Checklist Items", "Claude Insights", "Reports"])
        
        # Dashboard tab
        with tab1:
            if st.session_state.checklist_data is not None:
                st.header("Project Dashboard")
                
                # Key metrics in a row
                col1, col2, col3, col4 = st.columns(4)
                
                total_items = len(st.session_state.checklist_data)
                completed_items = len(st.session_state.checklist_data[st.session_state.checklist_data['status'].str.contains('Complete', case=False)])
                completion_percentage = (completed_items / total_items) * 100 if total_items > 0 else 0
                in_progress_items = len(st.session_state.checklist_data[st.session_state.checklist_data['status'].str.contains('Progress|Started', case=False)])
                not_started_items = len(st.session_state.checklist_data[st.session_state.checklist_data['status'].str.contains('Not Started', case=False)])
                high_priority_items = len(st.session_state.checklist_data[st.session_state.checklist_data['priority'] == 'High'])
                
                with col1:
                    st.metric("Total Items", total_items)
                
                with col2:
                    st.metric("Completion", f"{completion_percentage:.1f}%")
                
                with col3:
                    st.metric("In Progress", in_progress_items)
                
                with col4:
                    st.metric("High Priority", high_priority_items)
                
                # Main charts
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("Status Breakdown")
                    status_counts = st.session_state.checklist_data['status'].value_counts().reset_index()
                    status_counts.columns = ['Status', 'Count']
                    
                    fig = px.pie(status_counts, values='Count', names='Status', hole=0.4,
                                color_discrete_sequence=px.colors.qualitative.Set3)
                    fig.update_layout(height=300)
                    st.plotly_chart(fig, use_container_width=True)
                
                with col2:
                    st.subheader("Items by Category")
                    category_counts = st.session_state.checklist_data['category'].value_counts().reset_index()
                    category_counts.columns = ['Category', 'Count']
                    
                    fig = px.bar(category_counts, x='Category', y='Count', 
                                color='Count', color_continuous_scale='Viridis')
                    fig.update_layout(height=300)
                    st.plotly_chart(fig, use_container_width=True)
                
                # Completion by category chart
                st.subheader("Completion Status by Category")
                
                # Prepare data for stacked bar chart
                categories = st.session_state.checklist_data['category'].unique()
                status_data = []
                
                for cat in categories:
                    cat_data = st.session_state.checklist_data[st.session_state.checklist_data['category'] == cat]
                    cat_total = len(cat_data)
                    
                    # Count items by status
                    cat_complete = len(cat_data[cat_data['status'].str.contains('Complete', case=False)])
                    cat_in_progress = len(cat_data[cat_data['status'].str.contains('Progress|Started', case=False)])
                    cat_not_started = len(cat_data[cat_data['status'].str.contains('Not Started', case=False)])
                    
                    status_data.append({
                        'Category': cat,
                        'Complete': cat_complete,
                        'In Progress': cat_in_progress,
                        'Not Started': cat_not_started,
                        'Total': cat_total
                    })
                
                status_df = pd.DataFrame(status_data)
                
                # Create stacked bar chart
                fig = go.Figure()
                
                fig.add_trace(go.Bar(
                    x=status_df['Category'],
                    y=status_df['Complete'],
                    name='Complete',
                    marker_color='#00CC96'
                ))
                
                fig.add_trace(go.Bar(
                    x=status_df['Category'],
                    y=status_df['In Progress'],
                    name='In Progress',
                    marker_color='#FFA15A'
                ))
                
                fig.add_trace(go.Bar(
                    x=status_df['Category'],
                    y=status_df['Not Started'],
                    name='Not Started',
                    marker_color='#EF553B'
                ))
                
                fig.update_layout(
                    barmode='stack',
                    height=400,
                    margin=dict(l=20, r=20, t=30, b=20),
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # Claude insights on the dashboard
                if client and st.button("Generate Claude Insights"):
                    with st.spinner("Claude is analyzing your due diligence checklist..."):
                        st.session_state.claude_insights = analyze_checklist_with_claude(
                            client, 
                            st.session_state.checklist_data,
                            st.session_state.project_name
                        )
                        st.session_state.last_analyzed = datetime.now()
                    
                    st.success("Analysis complete!")
                
                # Display Claude insights if available
                if 'claude_insights' in st.session_state and st.session_state.claude_insights:
                    st.subheader("Claude AI Insights")
                    
                    if st.session_state.last_analyzed:
                        st.caption(f"Last analyzed: {st.session_state.last_analyzed.strftime('%Y-%m-%d %H:%M')}")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("#### Project Assessment")
                        st.info(st.session_state.claude_insights.get("OVERALL_ASSESSMENT", "No assessment available"))
                        
                        st.markdown("#### Key Risks")
                        risks = st.session_state.claude_insights.get("KEY_RISKS", ["No risks identified"])
                        for risk in risks:
                            st.warning(risk)
                    
                    with col2:
                        st.markdown("#### Suggestions")
                        suggestions = st.session_state.claude_insights.get("SUGGESTIONS", ["No suggestions available"])
                        for suggestion in suggestions:
                            st.success(suggestion)
                        
                        st.markdown("#### Next Steps")
                        next_steps = st.session_state.claude_insights.get("NEXT_STEPS", ["No next steps available"])
                        for step in next_steps:
                            st.info(step)
            else:
                st.info("Upload a due diligence checklist to get started")
        
        # Checklist Items tab
        with tab2:
            if st.session_state.filtered_data is not None:
                st.header("Due Diligence Checklist")
                
                # Add new item button
                if st.button("Add New Item"):
                    # Create new empty row
                    new_item = pd.DataFrame({
                        'item_description': ["New item - click to edit"],
                        'status': ["Not Started"],
                        'category': ["General"],
                        'owner': ["Unassigned"],
                        'due_date': [pd.Timestamp.now() + pd.Timedelta(days=14)],
                        'priority': ["Medium"],
                        'notes': [""],
                        'id': [str(uuid.uuid4())],
                        'last_updated': [pd.Timestamp.now()]
                    })
                    
                    # Append to existing data
                    st.session_state.checklist_data = pd.concat([st.session_state.checklist_data, new_item], ignore_index=True)
                    st.session_state.filtered_data = pd.concat([st.session_state.filtered_data, new_item], ignore_index=True)
                    st.success("New item added!")
                
                # Export button
                export_data = st.session_state.checklist_data.copy()
                export_buffer = export_to_excel(export_data)
                st.download_button(
                    label="Export Checklist",
                    data=export_buffer,
                    file_name=f"{st.session_state.project_name}_checklist_export.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Display editable checklist
                for index, row in st.session_state.filtered_data.iterrows():
                    with st.expander(f"{row['item_description']} ({row['status']})"):
                        # Create a form for each item
                        with st.form(key=f"item_form_{row['id']}"):
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                # Item description
                                new_description = st.text_area("Item Description", row['item_description'], height=100)
                                
                                # Status
                                new_status = st.selectbox(
                                    "Status",
                                    ["Not Started", "In Progress", "Complete", "Blocked", "Deferred"],
                                    index=["Not Started", "In Progress", "Complete", "Blocked", "Deferred"].index(row['status']) if row['status'] in ["Not Started", "In Progress", "Complete", "Blocked", "Deferred"] else 0
                                )
                                
                                # Category
                                new_category = st.text_input("Category", row['category'])
                                
                                # Owner
                                new_owner = st.text_input("Owner", row['owner'])
                            
                            with col2:
                                # Priority
                                new_priority = st.selectbox(
                                    "Priority",
                                    ["Low", "Medium", "High", "Critical"],
                                    index=["Low", "Medium", "High", "Critical"].index(row['priority']) if row['priority'] in ["Low", "Medium", "High", "Critical"] else 1
                                )
                                
                            # Due date
                                new_due_date = st.date_input(
                                    "Due Date",
                                    value=pd.to_datetime(row['due_date']).date() if pd.notna(row['due_date']) and isinstance(row['due_date'], pd.Timestamp) else datetime.now().date() + timedelta(days=14)
                                )
                                
                                # Notes
                                new_notes = st.text_area("Notes", row['notes'], height=100)
                            
                            # Save button
                            if st.form_submit_button("Save Changes"):
                                # Update the item in the dataframe
                                st.session_state.checklist_data.loc[st.session_state.checklist_data['id'] == row['id'], 'item_description'] = new_description
                                st.session_state.checklist_data.loc[st.session_state.checklist_data['id'] == row['id'], 'status'] = new_status
                                st.session_state.checklist_data.loc[st.session_state.checklist_data['id'] == row['id'], 'category'] = new_category
                                st.session_state.checklist_data.loc[st.session_state.checklist_data['id'] == row['id'], 'owner'] = new_owner
                                st.session_state.checklist_data.loc[st.session_state.checklist_data['id'] == row['id'], 'priority'] = new_priority
                                st.session_state.checklist_data.loc[st.session_state.checklist_data['id'] == row['id'], 'due_date'] = new_due_date
                                st.session_state.checklist_data.loc[st.session_state.checklist_data['id'] == row['id'], 'notes'] = new_notes
                                st.session_state.checklist_data.loc[st.session_state.checklist_data['id'] == row['id'], 'last_updated'] = pd.Timestamp.now()
                                
                                # Update filtered data as well
                                st.success("Item updated!")
                                
                                # Refresh the page to reflect changes
                                st.experimental_rerun()
                        
                        # Claude recommendations
                        if client:
                            if st.button("Get Claude's Recommendations", key=f"claude_rec_{row['id']}"):
                                with st.spinner("Getting recommendations..."):
                                    recommendations = get_claude_recommendations(client, row)
                                    st.info(recommendations)
            else:
                st.info("Upload a due diligence checklist to get started")
        
        # Claude Insights tab
        with tab3:
            st.header("Claude AI Insights")
            
            if client:
                if st.button("Analyze Checklist with Claude"):
                    with st.spinner("Claude is analyzing your due diligence checklist..."):
                        st.session_state.claude_insights = analyze_checklist_with_claude(
                            client, 
                            st.session_state.checklist_data,
                            st.session_state.project_name
                        )
                        st.session_state.last_analyzed = datetime.now()
                    
                    st.success("Analysis complete!")
                
                # Display comprehensive insights
                if 'claude_insights' in st.session_state and st.session_state.claude_insights:
                    if st.session_state.last_analyzed:
                        st.caption(f"Last analyzed: {st.session_state.last_analyzed.strftime('%Y-%m-%d %H:%M')}")
                    
                    # Overall Assessment
                    st.subheader("Overall Assessment")
                    st.info(st.session_state.claude_insights.get("OVERALL_ASSESSMENT", "No assessment available"))
                    
                    # Key Risks
                    st.subheader("Key Risks")
                    risks = st.session_state.claude_insights.get("KEY_RISKS", ["No risks identified"])
                    for risk in risks:
                        st.warning(risk)
                    
                    # Suggestions
                    st.subheader("Suggestions")
                    suggestions = st.session_state.claude_insights.get("SUGGESTIONS", ["No suggestions available"])
                    for suggestion in suggestions:
                        st.success(suggestion)
                    
                    # Category Analysis
                    st.subheader("Category Analysis")
                    st.info(st.session_state.claude_insights.get("CATEGORIES_ANALYSIS", "No category analysis available"))
                    
                    # Next Steps
                    st.subheader("Recommended Next Steps")
                    next_steps = st.session_state.claude_insights.get("NEXT_STEPS", ["No next steps available"])
                    for step in next_steps:
                        st.info(step)
                    
                    # Custom query
                    st.subheader("Ask Claude About Your Checklist")
                    user_query = st.text_area("Enter your question about the checklist", height=100)
                    
                    if user_query and st.button("Get Answer"):
                        with st.spinner("Claude is thinking..."):
                            try:
                                # Prepare context with checklist summary
                                checklist_summary = st.session_state.checklist_data[['item_description', 'category', 'status', 'priority']].head(50).to_string(index=False)
                                
                                prompt = f"""
                                You are analyzing a due diligence checklist for project: {st.session_state.project_name}.
                                
                                The user asks: {user_query}
                                
                                Here's a summary of the checklist (showing first 50 items):
                                {checklist_summary}
                                
                                Please provide a helpful, concise response to the user's question.
                                """
                                
                                message = client.messages.create(
                                    model="claude-3-haiku-20240307",
                                    max_tokens=1000,
                                    temperature=0.2,
                                    messages=[
                                        {"role": "user", "content": prompt}
                                    ]
                                )
                                
                                st.write(message.content[0].text)
                                
                            except Exception as e:
                                st.error(f"Error getting answer from Claude: {str(e)}")
                else:
                    st.info("Click 'Analyze Checklist with Claude' to get AI insights")
            else:
                st.warning("Claude API key not configured. Please add your API key in settings to enable AI insights.")
        
        # Reports tab
        with tab4:
            st.header("Reports & Exports")
            
            report_type = st.selectbox(
                "Report Type",
                ["Status Summary", "Category Analysis", "Owner Workload", "Upcoming Due Dates", "Full Report"]
            )
            
            if st.session_state.checklist_data is not None:
                if report_type == "Status Summary":
                    st.subheader("Status Summary Report")
                    
                    # Status breakdown
                    status_counts = st.session_state.checklist_data['status'].value_counts().reset_index()
                    status_counts.columns = ['Status', 'Count']
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Status pie chart
                        fig = px.pie(status_counts, values='Count', names='Status', 
                                    title="Items by Status",
                                    color_discrete_sequence=px.colors.qualitative.Set3)
                        st.plotly_chart(fig, use_container_width=True)
                    
                    with col2:
                        # Status table
                        st.dataframe(status_counts, use_container_width=True)
                        
                        # Completion percentage
                        total_items = len(st.session_state.checklist_data)
                        completed_items = len(st.session_state.checklist_data[st.session_state.checklist_data['status'].str.contains('Complete', case=False)])
                        completion_percentage = (completed_items / total_items) * 100 if total_items > 0 else 0
                        
                        st.metric("Overall Completion", f"{completion_percentage:.1f}%")
                    
                    # List of incomplete high priority items
                    st.subheader("High Priority Items Not Completed")
                    high_priority_incomplete = st.session_state.checklist_data[
                        (st.session_state.checklist_data['priority'] == 'High') & 
                        (~st.session_state.checklist_data['status'].str.contains('Complete', case=False))
                    ]
                    
                    if not high_priority_incomplete.empty:
                        st.dataframe(high_priority_incomplete[['item_description', 'status', 'category', 'owner', 'due_date']], use_container_width=True)
                    else:
                        st.success("No high priority incomplete items!")
                
                elif report_type == "Category Analysis":
                    st.subheader("Category Analysis Report")
                    
                    # Category breakdown
                    category_counts = st.session_state.checklist_data['category'].value_counts().reset_index()
                    category_counts.columns = ['Category', 'Count']
                    
                    # Calculate completion by category
                    categories = []
                    for category in st.session_state.checklist_data['category'].unique():
                        category_data = st.session_state.checklist_data[st.session_state.checklist_data['category'] == category]
                        total = len(category_data)
                        completed = len(category_data[category_data['status'].str.contains('Complete', case=False)])
                        completion_pct = (completed / total) * 100 if total > 0 else 0
                        
                        categories.append({
                            'Category': category,
                            'Total Items': total,
                            'Completed': completed,
                            'Completion %': completion_pct
                        })
                    
                    category_df = pd.DataFrame(categories)
                    
                    # Display category completion chart
                    fig = px.bar(category_df, x='Category', y='Completion %', 
                                 title="Completion Percentage by Category",
                                 color='Completion %', color_continuous_scale='RdYlGn')
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Display category data
                    st.dataframe(category_df, use_container_width=True)
                    
                    # Category with least progress
                    if not category_df.empty:
                        worst_category = category_df.loc[category_df['Completion %'].idxmin()]
                        st.warning(f"Category with least progress: **{worst_category['Category']}** ({worst_category['Completion %']:.1f}% complete)")
                        
                        # Show items in the least progressed category
                        st.subheader(f"Items in {worst_category['Category']}")
                        category_items = st.session_state.checklist_data[st.session_state.checklist_data['category'] == worst_category['Category']]
                        st.dataframe(category_items[['item_description', 'status', 'priority', 'owner', 'due_date']], use_container_width=True)
                
                elif report_type == "Owner Workload":
                    st.subheader("Owner Workload Report")
                    
                    # Items per owner
                    owner_counts = st.session_state.checklist_data['owner'].value_counts().reset_index()
                    owner_counts.columns = ['Owner', 'Total Items']
                    
                    # Calculate workload by owner
                    owners = []
                    for owner in st.session_state.checklist_data['owner'].unique():
                        owner_data = st.session_state.checklist_data[st.session_state.checklist_data['owner'] == owner]
                        total = len(owner_data)
                        completed = len(owner_data[owner_data['status'].str.contains('Complete', case=False)])
                        remaining = total - completed
                        completion_pct = (completed / total) * 100 if total > 0 else 0
                        high_priority = len(owner_data[owner_data['priority'].isin(['High', 'Critical'])])
                        
                        owners.append({
                            'Owner': owner,
                            'Total Items': total,
                            'Completed': completed,
                            'Remaining': remaining,
                            'Completion %': completion_pct,
                            'High Priority Items': high_priority
                        })
                    
                    owner_df = pd.DataFrame(owners)
                    
                    # Display owner workload chart
                    fig = px.bar(owner_df, x='Owner', y=['Completed', 'Remaining'], 
                                 title="Workload by Owner",
                                 barmode='stack', 
                                 color_discrete_map={'Completed': 'green', 'Remaining': 'red'})
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Display owner data
                    st.dataframe(owner_df.sort_values('Remaining', ascending=False), use_container_width=True)
                    
                    # Owner with highest workload
                    if not owner_df.empty:
                        highest_workload_owner = owner_df.loc[owner_df['Remaining'].idxmax()]
                        st.warning(f"Owner with highest remaining workload: **{highest_workload_owner['Owner']}** ({highest_workload_owner['Remaining']} items remaining)")
                        
                        # Show items for the highest workload owner
                        st.subheader(f"Items assigned to {highest_workload_owner['Owner']}")
                        owner_items = st.session_state.checklist_data[
                            (st.session_state.checklist_data['owner'] == highest_workload_owner['Owner']) & 
                            (~st.session_state.checklist_data['status'].str.contains('Complete', case=False))
                        ]
                        st.dataframe(owner_items[['item_description', 'status', 'category', 'priority', 'due_date']], use_container_width=True)
                
                elif report_type == "Upcoming Due Dates":
                    st.subheader("Upcoming Due Dates Report")
                    
                    # Calculate days until due
                    current_date = pd.Timestamp.now().date()
                    upcoming_df = st.session_state.checklist_data.copy()
                    
                    # Convert due_date to datetime if it's not already
                    if upcoming_df['due_date'].dtype != 'datetime64[ns]':
                        upcoming_df['due_date'] = pd.to_datetime(upcoming_df['due_date'])
                    
                    # Filter out items without due dates
                    upcoming_df = upcoming_df[upcoming_df['due_date'].notna()]
                    
                    # Add days_until_due column
                    upcoming_df['days_until_due'] = upcoming_df['due_date'].apply(
                        lambda x: (x.date() - current_date).days if isinstance(x, pd.Timestamp) else 0
                    )
                    
                    # Filter for incomplete items only
                    upcoming_df = upcoming_df[~upcoming_df['status'].str.contains('Complete', case=False)]
                    
                    # Sort by due date
                    upcoming_df = upcoming_df.sort_values('days_until_due')
                    
                    # Overdue items
                    overdue_items = upcoming_df[upcoming_df['days_until_due'] < 0]
                    if not overdue_items.empty:
                        st.error(f"**{len(overdue_items)} overdue items**")
                        st.dataframe(overdue_items[['item_description', 'days_until_due', 'status', 'category', 'priority', 'owner']], use_container_width=True)
                    
                    # Due this week
                    due_this_week = upcoming_df[(upcoming_df['days_until_due'] >= 0) & (upcoming_df['days_until_due'] <= 7)]
                    if not due_this_week.empty:
                        st.warning(f"**{len(due_this_week)} items due this week**")
                        st.dataframe(due_this_week[['item_description', 'days_until_due', 'status', 'category', 'priority', 'owner']], use_container_width=True)
                    
                    # Due next week
                    due_next_week = upcoming_df[(upcoming_df['days_until_due'] > 7) & (upcoming_df['days_until_due'] <= 14)]
                    if not due_next_week.empty:
                        st.info(f"**{len(due_next_week)} items due next week**")
                        st.dataframe(due_next_week[['item_description', 'days_until_due', 'status', 'category', 'priority', 'owner']], use_container_width=True)
                    
                    # Due later
                    due_later = upcoming_df[upcoming_df['days_until_due'] > 14]
                    if not due_later.empty:
                        st.success(f"**{len(due_later)} items due later**")
                        st.dataframe(due_later[['item_description', 'days_until_due', 'status', 'category', 'priority', 'owner']], use_container_width=True)
                
                elif report_type == "Full Report":
                    st.subheader("Full Due Diligence Report")
                    
                    # Project summary
                    st.markdown("### Project Summary")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    
                    total_items = len(st.session_state.checklist_data)
                    completed_items = len(st.session_state.checklist_data[st.session_state.checklist_data['status'].str.contains('Complete', case=False)])
                    completion_percentage = (completed_items / total_items) * 100 if total_items > 0 else 0
                    in_progress_items = len(st.session_state.checklist_data[st.session_state.checklist_data['status'].str.contains('Progress|Started', case=False)])
                    not_started_items = len(st.session_state.checklist_data[st.session_state.checklist_data['status'].str.contains('Not Started', case=False)])
                    
                    with col1:
                        st.metric("Total Items", total_items)
                    
                    with col2:
                        st.metric("Completion", f"{completion_percentage:.1f}%")
                    
                    with col3:
                        st.metric("In Progress", in_progress_items)
                    
                    with col4:
                        st.metric("Not Started", not_started_items)
                    
                    # Status chart
                    status_counts = st.session_state.checklist_data['status'].value_counts().reset_index()
                    status_counts.columns = ['Status', 'Count']
                    
                    fig = px.pie(status_counts, values='Count', names='Status', 
                                title="Items by Status",
                                color_discrete_sequence=px.colors.qualitative.Set3)
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Category completion
                    st.markdown("### Category Completion")
                    
                    # Calculate completion by category
                    categories = []
                    for category in st.session_state.checklist_data['category'].unique():
                        category_data = st.session_state.checklist_data[st.session_state.checklist_data['category'] == category]
                        total = len(category_data)
                        completed = len(category_data[category_data['status'].str.contains('Complete', case=False)])
                        completion_pct = (completed / total) * 100 if total > 0 else 0
                        
                        categories.append({
                            'Category': category,
                            'Total Items': total,
                            'Completed': completed,
                            'Completion %': completion_pct
                        })
                    
                    category_df = pd.DataFrame(categories)
                    
                    # Display category completion chart
                    fig = px.bar(category_df, x='Category', y='Completion %', 
                                 title="Completion Percentage by Category",
                                 color='Completion %', color_continuous_scale='RdYlGn')
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # High priority items
                    st.markdown("### High Priority Items")
                    high_priority_items = st.session_state.checklist_data[st.session_state.checklist_data['priority'] == 'High']
                    if not high_priority_items.empty:
                        st.dataframe(high_priority_items[['item_description', 'status', 'category', 'owner', 'due_date']], use_container_width=True)
                    else:
                        st.success("No high priority items!")
                    
                    # Upcoming due dates
                    st.markdown("### Upcoming Due Dates")
                    
                    # Calculate days until due
                    current_date = pd.Timestamp.now().date()
                    upcoming_df = st.session_state.checklist_data.copy()
                    
                    # Ensure due_date is properly formatted
                    upcoming_df = upcoming_df[upcoming_df['due_date'].notna()]
                    
                    # Safely calculate days until due
                    upcoming_df['days_until_due'] = upcoming_df['due_date'].apply(
                        lambda x: (x.date() - current_date).days if isinstance(x, pd.Timestamp) else 0
                    )
                    
                    # Filter for incomplete items only
                    upcoming_df = upcoming_df[~upcoming_df['status'].str.contains('Complete', case=False)]
                    
                    # Sort by due date
                    upcoming_df = upcoming_df.sort_values('days_until_due')
                    
                    # Display upcoming due dates
                    if not upcoming_df.empty:
                        st.dataframe(upcoming_df[['item_description', 'days_until_due', 'status', 'category', 'priority', 'owner']], use_container_width=True)
                    else:
                        st.success("No upcoming due dates!")
                    
                    # Claude insights
                    if 'claude_insights' in st.session_state and st.session_state.claude_insights:
                        st.markdown("### Claude AI Insights")
                        
                        # Overall Assessment
                        st.subheader("Overall Assessment")
                        st.info(st.session_state.claude_insights.get("OVERALL_ASSESSMENT", "No assessment available"))
                        
                        # Key Risks
                        st.subheader("Key Risks")
                        risks = st.session_state.claude_insights.get("KEY_RISKS", ["No risks identified"])
                        for risk in risks:
                            st.warning(risk)
                        
                        # Suggestions
                        st.subheader("Suggestions")
                        suggestions = st.session_state.claude_insights.get("SUGGESTIONS", ["No suggestions available"])
                        for suggestion in suggestions:
                            st.success(suggestion)
                        
                        # Next Steps
                        st.subheader("Recommended Next Steps")
                        next_steps = st.session_state.claude_insights.get("NEXT_STEPS", ["No next steps available"])
                        for step in next_steps:
                            st.info(step)
                
                # Export report button
                report_buffer = export_to_excel(st.session_state.checklist_data)
                st.download_button(
                    label=f"Export {report_type} Report",
                    data=report_buffer,
                    file_name=f"{st.session_state.project_name}_{report_type.lower().replace(' ', '_')}_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("Upload a due diligence checklist to generate reports")
    else:
        # Welcome screen
        st.title("Due Diligence Tracker")
        st.markdown("""
        ### Welcome to the Claude-powered Due Diligence Tracker
        
        This application helps you manage and track due diligence checklists for projects, with AI-powered insights from Claude.
        
        **Features:**
        - Import existing due diligence checklists from Excel
        - Track status of checklist items
        - Get AI-powered insights and recommendations
        - Generate reports and visualizations
        - Export data for sharing
        
        **Getting Started:**
        1. Enter a project name in the sidebar
        2. Upload your due diligence checklist Excel file
        3. Optional: Add your Claude API key for AI-powered insights
        
        **About:**
        This application is built with Streamlit and powered by Claude AI.
        """)
        
        # Sample Excel template
        st.info("Need a template? Download a sample due diligence checklist template to get started.")
        
        # Create a sample Excel file
        sample_data = {
            'Item Description': [
                'Review financial statements',
                'Verify legal compliance',
                'Assess IT infrastructure',
                'Review customer contracts',
                'Evaluate supply chain',
                'Analyze market position',
                'Check employee records',
                'Review intellectual property',
                'Assess environmental compliance',
                'Review tax records'
            ],
            'Status': ['Not Started'] * 10,
            'Category': [
                'Financial',
                'Legal',
                'IT',
                'Contracts',
                'Operations',
                'Market',
                'HR',
                'Legal',
                'Compliance',
                'Financial'
            ],
            'Owner': ['Unassigned'] * 10,
            'Due Date': [(datetime.now() + timedelta(days=14)).date()] * 10
        }
        
        sample_df = pd.DataFrame(sample_data)
        sample_buffer = export_to_excel(sample_df)
        
        st.download_button(
            label="Download Sample Template",
            data=sample_buffer,
            file_name="due_diligence_checklist_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Run the app
if __name__ == "__main__":
    main()
