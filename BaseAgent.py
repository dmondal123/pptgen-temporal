import json
import os
from temporalio.client import Client
from temporalio.worker import Worker
from BaseAgentWorkflow import PPTAgentWorkflow
from activities import (
    call_llm,
    create_memory_snapshot,
    execute_tool,
    extract_pptx_structure,
    extract_excel_structure,
    get_slide_xml,
    get_excel_table,
    modify_slide,
    modify_excel
)

def ensure_dir(file_path):
    """Ensure directory exists for given file path."""
    directory = os.path.dirname(file_path)
    if not os.path.exists(directory):
        os.makedirs(directory)

def write_json(file_path, data):
    """Write data to JSON file."""
    ensure_dir(file_path)
    with open(file_path, 'w') as json_file:
        json.dump(data, json_file, indent=4)
    print(f"Config written to {os.path.abspath(file_path)}")

def __add_context__(system_msg, files_info):
    """Add context to system message."""
    return f"""
    {system_msg}
    
    Available Files:
    {json.dumps(files_info, indent=2)}
    
    Instructions:
    1. Always examine file content before making modifications
    2. Use appropriate tools for viewing and modifying files
    3. Provide clear explanations of changes made
    """

class BaseAgent:
    def __init__(
        self,
        user_id="",
        system_msg="",
        pptx_files=None,
        excel_files=None,
        config_path="agent_configs"
    ):
        """Initialize the PowerPoint and Excel agent."""
        if pptx_files is None:
            pptx_files = []
        if excel_files is None:
            excel_files = []

        self.user_id = user_id
        self.pptx_files = pptx_files
        self.excel_files = excel_files
        
        # Create files info structure
        files_info = {
            "powerpoint_files": [os.path.basename(f) for f in pptx_files],
            "excel_files": [os.path.basename(f) for f in excel_files]
        }

        # Create agent configuration
        agent_config = {
            "system_msg": __add_context__(system_msg, files_info),
            "pptx_files": pptx_files,
            "excel_files": excel_files,
            "user_id": user_id
        }

        # Write configuration to file
        config_file = f"{config_path}/ppt_agent_{user_id}.json"
        write_json(config_file, agent_config)
        print(f"Agent initialized for user {user_id}")

    async def start_worker(self, interrupt_event):
        """Start the Temporal worker for this agent."""
        try:
            # Connect to Temporal server
            temporal_address = os.environ.get("TEMPORAL_ADDRESS", "localhost:7233")
            client = await Client.connect(temporal_address)

            # Create worker with all activities
            worker = Worker(
                client,
                task_queue=f"ppt-agent-{self.user_id}-queue",
                workflows=[PPTAgentWorkflow],
                activities=[
                    call_llm,
                    create_memory_snapshot,
                    execute_tool,
                    extract_pptx_structure,
                    extract_excel_structure,
                    get_slide_xml,
                    get_excel_table,
                    modify_slide,
                    modify_excel
                ]
            )

            print(f"Task queue: ppt-agent-{self.user_id}-queue")
            print("\nWorker started, ctrl+c to exit\n")

            # Run the worker until interrupted
            async with worker:
                await interrupt_event.wait()

        except Exception as e:
            print(f"Error in worker: {str(e)}")
        finally:
            print("\nShutting down the worker\n")

    async def start_workflow(self, workflow_id=None):
        """Start a new workflow instance."""
        try:
            temporal_address = os.environ.get("TEMPORAL_ADDRESS", "localhost:7233")
            client = await Client.connect(temporal_address)

            if workflow_id is None:
                workflow_id = f"ppt-agent-{self.user_id}-{os.urandom(4).hex()}"

            # Start the workflow
            handle = await client.start_workflow(
                PPTAgentWorkflow.run,
                id=workflow_id,
                task_queue=f"ppt-agent-{self.user_id}-queue"
            )

            print(f"Started workflow with ID: {workflow_id}")
            return handle

        except Exception as e:
            print(f"Error starting workflow: {str(e)}")
            return None

    async def send_user_query(self, workflow_handle, query: str):
        """Send a user query to the workflow."""
        try:
            # Prepare input data
            input_data = {
                "query": query,
                "pptx_files": self.pptx_files,
                "excel_files": self.excel_files
            }

            # Send signal to workflow
            await workflow_handle.signal(PPTAgentWorkflow.user_input, input_data)
            print(f"Sent query: {query}")

            # Wait briefly and get conversation history
            await asyncio.sleep(1)
            history = await workflow_handle.query(PPTAgentWorkflow.get_conversation_history)
            return history

        except Exception as e:
            print(f"Error sending query: {str(e)}")
            return None

async def run_agent(
    user_id: str,
    pptx_files: List[str],
    excel_files: List[str],
    system_msg: str = "I am an AI assistant that helps with PowerPoint and Excel files."
):
    """Helper function to run an agent instance."""
    import asyncio

    # Create and start agent
    agent = BaseAgent(
        user_id=user_id,
        system_msg=system_msg,
        pptx_files=pptx_files,
        excel_files=excel_files
    )

    # Create interrupt event
    interrupt_event = asyncio.Event()

    # Start worker in background task
    worker_task = asyncio.create_task(agent.start_worker(interrupt_event))

    try:
        # Start workflow
        workflow_handle = await agent.start_workflow()
        if workflow_handle is None:
            print("Failed to start workflow")
            return

        # Example: Send a test query
        history = await agent.send_user_query(
            workflow_handle,
            "What files are available?"
        )
        print("\nConversation history:")
        for msg in history:
            print(f"{msg['role']}: {msg['content'][:100]}...")

    except Exception as e:
        print(f"Error in run_agent: {str(e)}")
    finally:
        # Signal worker to stop
        interrupt_event.set()
        await worker_task

if __name__ == "__main__":
    import asyncio
    
    # Example usage
    test_files = {
        "pptx": ["example.pptx"],
        "excel": ["example.xlsx"]
    }
    
    asyncio.run(run_agent(
        user_id="test_user",
        pptx_files=test_files["pptx"],
        excel_files=test_files["excel"]
    ))
