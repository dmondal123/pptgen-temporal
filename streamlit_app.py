import streamlit as st
import asyncio
from temporalio.client import Client
import time
from datetime import datetime
import os
import uuid
import json

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
    if 'threads' not in st.session_state:
        st.session_state.threads = {}
    if 'current_thread' not in st.session_state:
        # Create first thread with a unique workflow ID
        thread_id = "Thread 1"
        workflow_id = f"ppt-agent-{thread_id}-{str(uuid.uuid4())[:8]}"
        st.session_state.threads[thread_id] = {
            "selected_pptx": [],
            "selected_excel": [],
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "workflow_id": workflow_id
        }
        st.session_state.current_thread = thread_id
    
    # Store files directory in session state
    if 'files_dir' not in st.session_state:
        st.session_state.files_dir = "files"
    
    # Track connection status
    if 'temporal_connected' not in st.session_state:
        st.session_state.temporal_connected = None
        
    # Track if we're waiting for a response
    if 'waiting_for_response' not in st.session_state:
        st.session_state.waiting_for_response = False
        
    # Track the count of assistant messages
    if 'assistant_msg_count' not in st.session_state:
        st.session_state.assistant_msg_count = 0
        
    # Store the latest user message
    if 'latest_user_message' not in st.session_state:
        st.session_state.latest_user_message = None

def create_new_thread():
    """Create a new thread with empty file selections and a unique workflow ID"""
    thread_id = f"Thread {len(st.session_state.threads) + 1}"
    workflow_id = f"ppt-agent-{thread_id}-{str(uuid.uuid4())[:8]}"
    
    st.session_state.threads[thread_id] = {
        "selected_pptx": [],
        "selected_excel": [],
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "workflow_id": workflow_id
    }
    
    st.session_state.current_thread = thread_id
    st.session_state.waiting_for_response = False
    st.session_state.latest_user_message = None
    return thread_id

async def get_temporal_client():
    """Connect to Temporal server with error handling"""
    try:
        client = await Client.connect("localhost:7233")
        st.session_state.temporal_connected = True
        return client
    except Exception as e:
        st.session_state.temporal_connected = False
        st.error(f"Could not connect to Temporal server: {str(e)}")
        return None

async def get_or_create_workflow(workflow_id):
    """Get or create workflow with specified ID"""
    client = await get_temporal_client()
    if not client:
        return None
    
    try:
        # Try to get existing workflow
        handle = client.get_workflow_handle(workflow_id)
        # Test if workflow is running
        await handle.query("get_conversation_history")
        return handle
    except Exception:
        # If workflow doesn't exist or isn't running, start a new one
        try:
            handle = await client.start_workflow(
                "PPTAgentWorkflow",
                id=workflow_id,
                task_queue="ppt-agent-task-queue"
            )
            return handle
        except Exception as e:
            st.error(f"Error creating workflow: {str(e)}")
            return None

async def get_conversation_history(workflow_id):
    """Get conversation history from the workflow"""
    try:
        client = await get_temporal_client()
        if not client:
            return []
        
        handle = client.get_workflow_handle(workflow_id)
        history = await handle.query("get_conversation_history")
        return history
    except Exception as e:
        st.warning(f"Could not fetch conversation history: {str(e)}")
        return []

async def send_user_input(workflow_id, query, pptx_files, excel_files):
    """Send user input to the workflow"""
    try:
        client = await get_temporal_client()
        if not client:
            return False
            
        handle = client.get_workflow_handle(workflow_id)
        input_data = {
            "query": query,
            "pptx_files": pptx_files,
            "excel_files": excel_files
        }
        await handle.signal("user_input", input_data)
        return True
    except Exception as e:
        st.error(f"Error sending message: {str(e)}")
        return False

def poll_for_assistant_response(workflow_id, message_count):
    """
    Poll until an assistant message with non-empty content is received,
    showing tool calls in real-time as they happen
    
    Args:
        workflow_id: ID of the workflow to query
        message_count: Count of assistant messages before sending user message
    
    Returns:
        True when a complete assistant response is detected
    """
    # Create placeholders for status elements
    message_placeholder = st.empty()
    progress_placeholder = st.empty()
    
    # Show waiting message with the info badge
    message_placeholder.info("Waiting for assistant response...")
    
    # Create a visual separator
    st.markdown("<hr style='margin-top: 0; margin-bottom: 15px; border-width: 1px; border-color: #f0f2f5;'>", unsafe_allow_html=True)
    
    # Create a container for displaying real-time updates AFTER the status elements
    update_container = st.container()
    
    max_attempts = 60
    delay = 1
    
    # Track which tool calls we've already seen
    seen_tool_call_ids = set()
    seen_tool_response_ids = set()
    
    for i in range(max_attempts):
        # Update progress bar
        progress_placeholder.progress((i + 1) / max_attempts)
        
        # Get current conversation history
        history = asyncio.run(get_conversation_history(workflow_id))
        
        # Find all assistant messages
        assistant_messages = [msg for msg in history if msg.get("role") == "assistant"]
        
        # Check if we have new tool calls to display
        if assistant_messages and len(assistant_messages) >= message_count:
            latest_assistant = assistant_messages[-1]
            tool_calls = latest_assistant.get("tool_calls", [])
            
            # Display any new tool calls
            if tool_calls:
                with update_container:
                    for tool_call in tool_calls:
                        tool_id = tool_call.get("id", "")
                        
                        # Only show tool calls we haven't seen before
                        if tool_id not in seen_tool_call_ids:
                            seen_tool_call_ids.add(tool_id)
                            
                            tool_name = tool_call["function"]["name"]
                            tool_args = tool_call["function"]["arguments"]
                            
                            # Create a GitHub-inspired tool call box
                            st.markdown(f"""
                            <div class="tool-call">
                                <div class="tool-call-header">
                                    <span class="tool-name">📦 {tool_name}</span>
                                </div>
                                <div class="tool-call-body">
                            """, unsafe_allow_html=True)
                            
                            # Try to format args as JSON
                            try:
                                args_dict = json.loads(tool_args)
                                formatted_args = json.dumps(args_dict, indent=2)
                            except:
                                formatted_args = tool_args
                                
                            st.code(formatted_args, language="json")
                            st.markdown("</div></div>", unsafe_allow_html=True)
            
            # Find new tool responses
            for tool_msg in history:
                if tool_msg.get("role") == "tool":
                    tool_call_id = tool_msg.get("tool_call_id", "")
                    
                    # Check if this is a response to a tool call we've seen but haven't seen the response yet
                    if tool_call_id in seen_tool_call_ids and tool_call_id not in seen_tool_response_ids:
                        seen_tool_response_ids.add(tool_call_id)
                        
                        # Get the tool name from the matching tool call
                        tool_name = "Unknown Tool"
                        for message in history:
                            if message.get("role") == "assistant" and message.get("tool_calls"):
                                for tc in message.get("tool_calls", []):
                                    if tc.get("id") == tool_call_id:
                                        tool_name = tc["function"]["name"]
                                        break
                        
                        with update_container:
                            # Get the tool response content
                            tool_response = tool_msg.get("content", "")
                            
                            # Display the tool response in a stylized box
                            st.markdown(f"""
                            <div class="tool-call">
                                <div class="tool-call-header" style="background-color: #e6ffed;">
                                    <span class="tool-response-header">✅ Response from: {tool_name}</span>
                                </div>
                                <div class="tool-call-body">
                            """, unsafe_allow_html=True)
                            
                            # Determine how to display the response content
                            if tool_response.startswith("<slide>") or tool_response.startswith("Error:"):
                                st.code(tool_response, language="xml")
                            elif "```" in tool_response or tool_response.startswith("|"):
                                # This is likely markdown or a table
                                st.markdown(tool_response)
                            else:
                                st.write(tool_response)
                                
                            st.markdown("</div></div>", unsafe_allow_html=True)
                            
                            # Add space between responses
                            st.markdown("<br>", unsafe_allow_html=True)
        
        # Check if we have a complete assistant response (non-empty content)
        if (len(assistant_messages) > message_count and 
            assistant_messages[-1].get("content", "") != ""):
            # Clear the status elements
            message_placeholder.empty()
            progress_placeholder.empty()
            return True
            
        # Wait before polling again
        time.sleep(delay)
        
        # Force a UI refresh (experimental)
        # This uses a trick to get Streamlit to refresh the UI during polling
        with st.empty():
            pass
    
    # Timeout occurred
    message_placeholder.warning("Timed out waiting for assistant response.")
    progress_placeholder.empty()
    return False

def list_files(directory, extensions):
    """List files with specific extensions in a directory"""
    files = []
    if os.path.exists(directory):
        for file in os.listdir(directory):
            if any(file.lower().endswith(ext) for ext in extensions):
                files.append(os.path.join(directory, file))
    return files

def save_uploaded_file(uploaded_file, directory):
    """Save an uploaded file to the specified directory and return the full path"""
    # Create directory if it doesn't exist
    if not os.path.exists(directory):
        os.makedirs(directory)
    
    # Save the file
    file_path = os.path.join(directory, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path

def count_assistant_messages(conversation):
    """Count how many messages from the assistant are in the conversation"""
    return len([msg for msg in conversation if msg.get("role") == "assistant"])

def display_conversation(conversation):
    """
    Display conversation with tool calls in a GitHub workspace-inspired UI
    with yellow tool names and green responses
    
    Args:
        conversation: List of message objects from the conversation history
    """
    # Skip the system message
    for i, message in enumerate(conversation[1:] if len(conversation) > 0 else []):
        role = message.get("role", "")
        content = message.get("content", "")
        
        # Handle user messages
        if role == "user":
            with st.chat_message(role):
                st.write(content)
                
        # Handle assistant messages with special handling for tool calls
        elif role == "assistant":
            with st.chat_message(role):
                # Display the text content
                if content:
                    st.write(content)
                
                # Check if there are tool calls
                tool_calls = message.get("tool_calls", [])
                if tool_calls:
                    # Create an expander for tool calls
                    with st.expander("Tool Calls", expanded=True):
                        # Display each tool call
                        for tool_call in tool_calls:
                            tool_id = tool_call.get("id", "")
                            tool_name = tool_call["function"]["name"]
                            tool_args = tool_call["function"]["arguments"]
                            
                            # Format JSON args for better display
                            try:
                                args_dict = json.loads(tool_args)
                                formatted_args = json.dumps(args_dict, indent=2)
                            except:
                                formatted_args = tool_args
                            
                            # Show the tool name in yellow
                            st.markdown(f'<div class="tool-name">📦 {tool_name}</div>', unsafe_allow_html=True)
                            
                            # Show the arguments in a code block
                            st.code(formatted_args, language="json")
                            
                            # Find corresponding tool response if available
                            for next_msg in conversation[i+2:]:  # +2 to account for system message and current index
                                if next_msg.get("role") == "tool" and next_msg.get("tool_call_id") == tool_id:
                                    tool_response = next_msg.get("content", "")
                                    
                                    # Display the response header in green
                                    st.markdown('<div class="tool-response-header">✅ Response</div>', unsafe_allow_html=True)
                                    
                                    # Determine how to display the response content
                                    if tool_response.startswith("<slide>") or tool_response.startswith("Error:"):
                                        st.code(tool_response, language="xml")
                                    elif "```" in tool_response or tool_response.startswith("|"):
                                        # This is likely markdown or a table
                                        st.markdown(tool_response)
                                    else:
                                        st.write(tool_response)
                                    break
                            
                            # Add some space between tool calls
                            st.markdown("<hr style='margin: 15px 0; opacity: 0.3;'>", unsafe_allow_html=True)

def main():
    # Initialize session state
    init_session_state()
    
    # Main layout
    st.sidebar.title("PowerPoint & Excel Assistant")
    
    # Files directory and extensions
    files_dir = st.session_state.files_dir
    pptx_extensions = [".pptx", ".ppt"]
    excel_extensions = [".xlsx", ".xls"]
    
    # Thread management and file selection in sidebar
    with st.sidebar:
        # Thread section
        col1, col2 = st.columns([6, 1])
        with col1:
            st.markdown("<h3>Threads</h3>", unsafe_allow_html=True)
        with col2:
            if st.button("➕", help="Create a new thread"):
                create_new_thread()
                st.rerun()
        
        # Thread selection
        for thread_id in st.session_state.threads:
            thread_data = st.session_state.threads[thread_id]
            # Create a styled thread box
            active_class = "active" if thread_id == st.session_state.current_thread else ""
            if st.button(
                f"{thread_id}",
                key=f"thread_{thread_id}",
                help=f"Created: {thread_data['created_at']}",
                use_container_width=True
            ):
                st.session_state.current_thread = thread_id
                st.session_state.waiting_for_response = False
                st.session_state.latest_user_message = None
                st.rerun()
        
        # Get current thread data
        current_thread = st.session_state.current_thread
        current_thread_data = st.session_state.threads[current_thread]
        workflow_id = current_thread_data["workflow_id"]
        
        # File selection for current thread
        st.subheader("File Selection")
        
        # File upload section
        with st.expander("Upload New Files", expanded=False):
            # Single file uploader for both PowerPoint and Excel
            uploaded_file = st.file_uploader(
                "Upload PowerPoint or Excel File",
                type=["ppt", "pptx", "xls", "xlsx"],
                key=f"{current_thread}_file_uploader",
                label_visibility="hidden"
            )
            if uploaded_file:
                # Save the uploaded file
                file_path = save_uploaded_file(uploaded_file, files_dir)
                
                # Determine file type based on extension and add to appropriate list
                file_name = uploaded_file.name.lower()
                if file_name.endswith((".ppt", ".pptx")):
                    # Add to PowerPoint list if not already selected
                    if file_path not in current_thread_data["selected_pptx"]:
                        current_thread_data["selected_pptx"].append(file_path)
                    st.success(f"Uploaded and selected PowerPoint: {uploaded_file.name}")
                elif file_name.endswith((".xls", ".xlsx")):
                    # Add to Excel list if not already selected
                    if file_path not in current_thread_data["selected_excel"]:
                        current_thread_data["selected_excel"].append(file_path)
                    st.success(f"Uploaded and selected Excel: {uploaded_file.name}")
        
        # List available files
        pptx_files = list_files(files_dir, pptx_extensions)
        excel_files = list_files(files_dir, excel_extensions)
        
        # PowerPoint files
        with st.expander("PowerPoint Files", expanded=True):
            selected_pptx = []
            for pptx_file in pptx_files:
                filename = os.path.basename(pptx_file)
                is_selected = pptx_file in current_thread_data["selected_pptx"]
                if st.checkbox(filename, value=is_selected, key=f"{current_thread}_pptx_{filename}"):
                    selected_pptx.append(pptx_file)
            
            # Update selected files
            current_thread_data["selected_pptx"] = selected_pptx
        
        # Excel files
        with st.expander("Excel Files", expanded=True):
            selected_excel = []
            for excel_file in excel_files:
                filename = os.path.basename(excel_file)
                is_selected = excel_file in current_thread_data["selected_excel"]
                if st.checkbox(filename, value=is_selected, key=f"{current_thread}_excel_{filename}"):
                    selected_excel.append(excel_file)
            
            # Update selected files
            current_thread_data["selected_excel"] = selected_excel
        
        # Connection status
        if st.session_state.temporal_connected is False:
            st.error("Disconnected from Temporal server")
        elif st.session_state.temporal_connected is True:
            st.success("Connected to Temporal server")
    
    # Chat area (now uses full width)
    # Display current thread info
    st.subheader(f"{current_thread}")
    
    # Show selected files
    selected_pptx = current_thread_data["selected_pptx"]
    selected_excel = current_thread_data["selected_excel"]
    
    if selected_pptx or selected_excel:
        files_text = []
        if selected_pptx:
            files_text.append(f"**PowerPoint:** {', '.join([os.path.basename(f) for f in selected_pptx])}")
        if selected_excel:
            files_text.append(f"**Excel:** {', '.join([os.path.basename(f) for f in selected_excel])}")
        
        st.info("Selected files: " + " | ".join(files_text))
    
    # Create or get workflow for this thread
    asyncio.run(get_or_create_workflow(workflow_id))
    
    # Get conversation history
    conversation = asyncio.run(get_conversation_history(workflow_id))
    
    # Check if we're waiting for a response or ready for input
    if st.session_state.waiting_for_response:
        # We're already polling, continue with that
        assistant_msg_count = st.session_state.assistant_msg_count
        
        # Display the conversation history 
        if conversation:
            display_conversation(conversation)
            
        # Also explicitly display the latest user message if it's not in the conversation yet
        # This ensures it's visible during streaming even if not yet included in the workflow history
        latest_user_prompt = st.session_state.latest_user_message
        if latest_user_prompt:
            # Check if the latest message is already in the conversation
            user_messages = [msg.get("content", "") for msg in conversation if msg.get("role") == "user"]
            if not user_messages or user_messages[-1] != latest_user_prompt:
                with st.chat_message("user"):
                    st.write(latest_user_prompt)
        
        # Continue polling
        if poll_for_assistant_response(workflow_id, assistant_msg_count):
            # Got full response, return to input mode
            st.session_state.waiting_for_response = False
            st.session_state.latest_user_message = None  # Clear the stored message
            st.rerun()
    else:
        # Display conversation history for non-polling mode
        if conversation:
            # Use our function to display the conversation
            display_conversation(conversation)
        else:
            st.info("Send a message to start the conversation.")
        
        # Chat input
        if prompt := st.chat_input("Ask about your PowerPoint or Excel files..."):
            # Count assistant messages before sending
            assistant_msg_count = count_assistant_messages(conversation)
            
            # Store in session state for polling
            st.session_state.assistant_msg_count = assistant_msg_count
            st.session_state.latest_user_message = prompt  # Store the latest user message
            
            # Show user message immediately
            with st.chat_message("user"):
                st.write(prompt)
            
            # Send message to the workflow
            success = asyncio.run(send_user_input(workflow_id, prompt, selected_pptx, selected_excel))
            
            if success:
                # Start polling and mark that we're waiting for a response
                st.session_state.waiting_for_response = True
                st.rerun()  # Rerun to enter polling mode
            else:
                if st.session_state.temporal_connected is False:
                    st.error("Message not sent: Disconnected from Temporal server")

if __name__ == "__main__":
    main() 
