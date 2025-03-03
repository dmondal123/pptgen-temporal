import streamlit as st
import asyncio
from temporalio.client import Client
import time
from datetime import datetime
import os
import uuid
import json
from BaseAgent import BaseAgent
from activities import (
    LLMParams,
    MemorySnapshotParams,
    ToolExecutionParams
)

# Set page config first
st.set_page_config(layout="wide", page_title="PowerPoint & Excel Assistant")

# Enhanced styling with GitHub workspace inspiration
st.markdown("""
<style>
/* Thread styling */
.thread-item {
    background-color: #f0f2f6;
    border-radius: 5px;
    padding: 10px;
    margin-bottom: 5px;
}
.thread-item.active {
    background-color: #e0e5ea;
    border-left: 5px solid #4e8cff;
}

/* GitHub-inspired styling */
.tool-call {
    background-color: #f6f8fa;
    border: 1px solid #e1e4e8;
    border-radius: 6px;
    margin-bottom: 16px;
    overflow: hidden;
}
.tool-call-header {
    background-color: #f1f8ff;
    border-bottom: 1px solid #e1e4e8;
    padding: 8px 16px;
    font-family: ui-monospace, SFMono-Regular, SF Mono, Menlo, Consolas, Liberation Mono, monospace;
}
.tool-call-body {
    padding: 16px;
}
.tool-response {
    background-color: #f8f9fa;
    border-top: 1px solid #e1e4e8;
    padding: 8px 16px;
}

/* Tool name and response styling */
.tool-name {
    color: #e9b914;
    font-weight: bold;
    font-family: monospace;
    background-color: #fffbe6;
    padding: 2px 6px;
    border-radius: 3px;
}
.tool-response-header {
    color: #28a745;
    font-weight: bold;
    background-color: #e6ffed;
    padding: 2px 6px;
    border-radius: 3px;
    margin-top: 10px;
}

/* File selection styling */
.file-section {
    margin-top: 20px;
}
</style>
""", unsafe_allow_html=True)

def init_session_state():
    """Initialize session state variables"""
    if 'agent' not in st.session_state:
        st.session_state.agent = None
    if 'threads' not in st.session_state:
        st.session_state.threads = {}
    if 'current_thread' not in st.session_state:
        thread_id = "Thread 1"
        workflow_id = f"ppt-agent-{thread_id}-{str(uuid.uuid4())[:8]}"
        st.session_state.threads[thread_id] = {
            "selected_pptx": [],
            "selected_excel": [],
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "workflow_id": workflow_id,
            "workflow_handle": None
        }
        st.session_state.current_thread = thread_id
    
    if 'files_dir' not in st.session_state:
        st.session_state.files_dir = "files"
    if 'temporal_connected' not in st.session_state:
        st.session_state.temporal_connected = None
    if 'waiting_for_response' not in st.session_state:
        st.session_state.waiting_for_response = False

async def initialize_agent(user_id: str, pptx_files: list, excel_files: list):
    """Initialize or get the agent for the current session"""
    if not st.session_state.agent:
        system_msg = """I am an AI assistant that helps with PowerPoint and Excel files.
        I can examine and modify slides and spreadsheets based on your requests."""
        
        st.session_state.agent = BaseAgent(
            user_id=user_id,
            system_msg=system_msg,
            pptx_files=pptx_files,
            excel_files=excel_files
        )
    return st.session_state.agent

async def get_or_create_workflow(thread_data):
    """Get existing workflow handle or create new workflow"""
    try:
        if not thread_data.get("workflow_handle"):
            agent = await initialize_agent(
                thread_data["workflow_id"],
                thread_data["selected_pptx"],
                thread_data["selected_excel"]
            )
            workflow_handle = await agent.start_workflow(thread_data["workflow_id"])
            thread_data["workflow_handle"] = workflow_handle
            st.session_state.temporal_connected = True
        return thread_data["workflow_handle"]
    except Exception as e:
        st.session_state.temporal_connected = False
        print(f"Error connecting to workflow: {str(e)}")
        return None

async def send_user_input(workflow_handle, prompt: str, pptx_files: list, excel_files: list):
    """Send user input to the workflow"""
    try:
        if not workflow_handle:
            return False
        
        history = await st.session_state.agent.send_user_query(
            workflow_handle,
            prompt
        )
        return True
    except Exception as e:
        st.session_state.temporal_connected = False
        print(f"Error sending user input: {str(e)}")
        return False

def display_conversation(conversation):
    """Display the conversation history"""
    for message in conversation:
        role = message.get("role", "")
        content = message.get("content", "")
        
        if role == "user":
            with st.chat_message("user"):
                st.write(content)
        
        elif role == "assistant":
            with st.chat_message("assistant"):
                st.write(content)
                
                # Display tool calls if present
                tool_calls = message.get("tool_calls", [])
                if tool_calls:
                    for tool_call in tool_calls:
                        tool_name = tool_call["function"]["name"]
                        tool_args = tool_call["function"]["arguments"]
                        
                        try:
                            args_dict = json.loads(tool_args)
                            formatted_args = json.dumps(args_dict, indent=2)
                        except:
                            formatted_args = tool_args
                        
                        st.markdown(f'<div class="tool-name">📦 {tool_name}</div>', unsafe_allow_html=True)
                        st.code(formatted_args, language="json")
        
        elif role == "tool":
            tool_name = message.get("name", "")
            tool_response = message.get("content", "")
            
            st.markdown(f'<div class="tool-response-header">✅ Response from {tool_name}</div>', unsafe_allow_html=True)
            
            if tool_response.startswith("<slide>") or tool_response.startswith("Error:"):
                st.code(tool_response, language="xml")
            elif "```" in tool_response or tool_response.startswith("|"):
                st.markdown(tool_response)
            else:
                st.write(tool_response)

def main():
    # Initialize session state
    init_session_state()
    
    # Main layout
    st.sidebar.title("PowerPoint & Excel Assistant")
    
    # Files directory
    files_dir = st.session_state.files_dir
    
    # Thread management in sidebar
    with st.sidebar:
        col1, col2 = st.columns([6, 1])
        with col1:
            st.markdown("<h3>Threads</h3>", unsafe_allow_html=True)
        with col2:
            if st.button("➕", help="Create new thread"):
                thread_id = f"Thread {len(st.session_state.threads) + 1}"
                workflow_id = f"ppt-agent-{thread_id}-{str(uuid.uuid4())[:8]}"
                st.session_state.threads[thread_id] = {
                    "selected_pptx": [],
                    "selected_excel": [],
                    "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "workflow_id": workflow_id,
                    "workflow_handle": None
                }
                st.session_state.current_thread = thread_id
                st.rerun()
        
        # Thread selection
        for thread_id, thread_data in st.session_state.threads.items():
            if st.button(
                f"{thread_id}",
                key=f"thread_{thread_id}",
                help=f"Created: {thread_data['created_at']}",
                use_container_width=True
            ):
                st.session_state.current_thread = thread_id
                st.session_state.waiting_for_response = False
                st.rerun()
        
        # File management for current thread
        current_thread = st.session_state.current_thread
        current_thread_data = st.session_state.threads[current_thread]
        
        # File upload and selection UI
        st.subheader("File Selection")
        
        # File upload
        with st.expander("Upload Files", expanded=False):
            uploaded_file = st.file_uploader(
                "Upload PowerPoint or Excel File",
                type=["ppt", "pptx", "xls", "xlsx"],
                key=f"{current_thread}_uploader"
            )
            
            if uploaded_file:
                file_path = os.path.join(files_dir, uploaded_file.name)
                os.makedirs(files_dir, exist_ok=True)
                
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                if uploaded_file.name.lower().endswith((".ppt", ".pptx")):
                    if file_path not in current_thread_data["selected_pptx"]:
                        current_thread_data["selected_pptx"].append(file_path)
                    st.success(f"Added PowerPoint: {uploaded_file.name}")
                else:
                    if file_path not in current_thread_data["selected_excel"]:
                        current_thread_data["selected_excel"].append(file_path)
                    st.success(f"Added Excel: {uploaded_file.name}")
        
        # File selection
        with st.expander("PowerPoint Files", expanded=True):
            pptx_files = [f for f in os.listdir(files_dir) if f.lower().endswith((".ppt", ".pptx"))]
            selected_pptx = []
            for pptx_file in pptx_files:
                file_path = os.path.join(files_dir, pptx_file)
                if st.checkbox(pptx_file, value=file_path in current_thread_data["selected_pptx"]):
                    selected_pptx.append(file_path)
            current_thread_data["selected_pptx"] = selected_pptx
        
        with st.expander("Excel Files", expanded=True):
            excel_files = [f for f in os.listdir(files_dir) if f.lower().endswith((".xls", ".xlsx"))]
            selected_excel = []
            for excel_file in excel_files:
                file_path = os.path.join(files_dir, excel_file)
                if st.checkbox(excel_file, value=file_path in current_thread_data["selected_excel"]):
                    selected_excel.append(file_path)
            current_thread_data["selected_excel"] = selected_excel
    
    # Main chat area
    st.subheader(current_thread)
    
    # Show selected files
    if current_thread_data["selected_pptx"] or current_thread_data["selected_excel"]:
        files_text = []
        if current_thread_data["selected_pptx"]:
            files_text.append(f"**PowerPoint:** {', '.join(os.path.basename(f) for f in current_thread_data['selected_pptx'])}")
        if current_thread_data["selected_excel"]:
            files_text.append(f"**Excel:** {', '.join(os.path.basename(f) for f in current_thread_data['selected_excel'])}")
        st.info("Selected files: " + " | ".join(files_text))
    
    # Get or create workflow
    workflow_handle = asyncio.run(get_or_create_workflow(current_thread_data))
    
    if workflow_handle:
        # Get conversation history
        history = asyncio.run(workflow_handle.query("get_conversation_history"))
        
        # Display conversation
        if history:
            display_conversation(history)
        else:
            st.info("Send a message to start the conversation.")
        
        # Chat input
        if prompt := st.chat_input("Ask about your files..."):
            with st.chat_message("user"):
                st.write(prompt)
            
            success = asyncio.run(send_user_input(
                workflow_handle,
                prompt,
                current_thread_data["selected_pptx"],
                current_thread_data["selected_excel"]
            ))
            
            if success:
                st.rerun()
            else:
                st.error("Failed to send message")
    else:
        st.error("Not connected to workflow server")

if __name__ == "__main__":
    main() 
