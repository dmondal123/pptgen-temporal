import sys
import asyncio
from slides_agent.BaseAgent import BaseAgent
from dataclasses import dataclass
from typing import List

@dataclass
class FileConfig:
    pptx_files: List[str]
    excel_files: List[str]

interrupt_event = asyncio.Event()

if __name__ == "__main__":
    from dotenv import load_dotenv
    load_dotenv()

    # Get user ID from command line arguments
    if len(sys.argv) < 2:
        print("Usage: python worker.py <user_id>")
        sys.exit(1)

    user_id = sys.argv[1]
    loop = asyncio.new_event_loop()

    # Define file paths for testing
    file_config = FileConfig(
        pptx_files=[
            "presentations/sample.pptx",
            "presentations/template.pptx"
        ],
        excel_files=[
            "spreadsheets/data.xlsx",
            "spreadsheets/report.xlsx"
        ]
    )

    system_message = f"""
    You are a dedicated PowerPoint and Excel assistant for {user_id}.
    Your capabilities include:
    
    PowerPoint:
    - Examining slide content and structure
    - Making modifications to slides
    - Adding or updating text, shapes, and tables
    - Formatting and styling elements
    
    Excel:
    - Reading sheet data and structure
    - Modifying cell contents and formulas
    - Creating and updating tables
    - Data analysis and formatting
    
    Guidelines:
    - Always examine file content before making modifications
    - Provide clear explanations of changes made
    - Use appropriate tools for viewing and modifying files
    - Confirm successful modifications
    - Handle errors gracefully and inform the user
    
    Remember to:
    - Be precise in your modifications
    - Keep track of file versions
    - Explain your actions clearly
    - Ask for clarification when needed
    """

    agent = BaseAgent(
        user_id=user_id,
        system_msg=system_message,
        pptx_files=file_config.pptx_files,
        excel_files=file_config.excel_files
    )

    try:
        # Start the worker
        print(f"Starting worker for user {user_id}...")
        loop.run_until_complete(agent.start_worker(interrupt_event=interrupt_event))
    except KeyboardInterrupt:
        print("\nInterrupt received, shutting down...")
        interrupt_event.set()
        loop.run_until_complete(loop.shutdown_asyncgens())
    except Exception as e:
        print(f"Error running worker: {str(e)}")
        interrupt_event.set()
    finally:
        loop.close()
