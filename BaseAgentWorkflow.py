import _io
import json
from typing import Union, List, Dict, Any
from temporalio import workflow
from datetime import timedelta
import asyncio
import os

from activities import (
    LLMParams,
    MemorySnapshotParams,
    ToolExecutionParams,
    call_llm,
    create_memory_snapshot,
    execute_tool
)

# Define tools for the LLM
def define_tools():
    return [
        {
            "type": "function",
            "function": {
                "name": "get_slide",
                "description": "Get the XML representation of a slide from a PowerPoint file",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the PowerPoint file"
                        },
                        "slide_index": {
                            "type": "integer",
                            "description": "Zero-based index of the slide to retrieve"
                        }
                    },
                    "required": ["file_path", "slide_index"]
                }
            }
        },
        {
            "type": "function",
            "function": {
                "name": "get_excel_data",
                "description": "Get the data from an Excel sheet as a markdown table",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the Excel file"
                        },
                        "sheet_name": {
                            "type": "string",
                            "description": "Name of the sheet to retrieve"
                        }
                    },
                    "required": ["file_path", "sheet_name"]
                }
            }
        },
        {
            "type": "function",
            "function": {
                "name": "modify_slide",
                "description": "Modify a slide using Python code",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the PowerPoint file"
                        },
                        "slide_index": {
                            "type": "integer",
                            "description": "Zero-based index of the slide to modify"
                        },
                        "code": {
                            "type": "string",
                            "description": "Python code to execute to modify the slide (has access to 'slide' object from python-pptx)"
                        }
                    },
                    "required": ["file_path", "slide_index", "code"]
                }
            }
        },
        {
            "type": "function",
            "function": {
                "name": "modify_excel",
                "description": "Modify an Excel sheet using Python code",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the Excel file"
                        },
                        "sheet_name": {
                            "type": "string",
                            "description": "Name of the sheet to modify"
                        },
                        "code": {
                            "type": "string",
                            "description": "Python code to execute to modify the sheet (has access to 'df' DataFrame from pandas)"
                        }
                    },
                    "required": ["file_path", "sheet_name", "code"]
                }
            }
        }
    ]

@workflow.defn
class PPTAgentWorkflow:
    def __init__(self):
        self.messages = []
        self.pptx_files = []
        self.excel_files = []
        self.file_path_mapping = {}
        self.memory = {}
        self.tools = define_tools()
        self.user_input_received = False
        self.user_query = ""

    @workflow.run
    async def run(self) -> List[Dict[str, Any]]:
        """Main workflow that orchestrates the agent."""
        # Initialize with system message
        self.messages = [
            {
                "role": "system",
                "content": "You are an AI PowerPoint and Excel agent. You can view and modify PowerPoint slides and Excel sheets."
            }
        ]

        while True:
            # Wait for user input
            await workflow.wait_condition(lambda: bool(self.user_input_received))
            self.user_input_received = False

            # Update memory snapshot
            memory_params = MemorySnapshotParams(
                pptx_files=self.pptx_files,
                excel_files=self.excel_files
            )
            self.memory = await workflow.execute_activity(
                create_memory_snapshot,
                args=[memory_params],
                start_to_close_timeout=timedelta(seconds=30)
            )

            # Update system message with memory
            memory_str = json.dumps(self.memory, indent=2)
            file_mapping_str = json.dumps(self.file_path_mapping, indent=2)
            
            self.messages[0]["content"] = f"""You are an AI PowerPoint and Excel agent. You can view and modify PowerPoint slides and Excel sheets.

Available files:
{memory_str}

File paths:
{file_mapping_str}

You have access to tools to:
1. View slide content (get_slide)
2. View Excel data (get_excel_data)
3. Modify slides (modify_slide)
4. Modify Excel sheets (modify_excel)

Always examine file content before making modifications."""

            # Add user query to messages
            self.messages.append({
                "role": "user",
                "content": self.user_query
            })

            # Process message chain until LLM stops calling tools
            while True:
                # Call LLM
                llm_params = LLMParams(
                    messages=self.messages,
                    tools=self.tools
                )
                
                assistant_message = await workflow.execute_activity(
                    call_llm,
                    args=[llm_params],
                    start_to_close_timeout=timedelta(minutes=2)
                )

                # Add assistant message to conversation
                self.messages.append({
                    "role": "assistant",
                    "content": assistant_message["content"],
                    "tool_calls": assistant_message["tool_calls"]
                })

                # Check if tools need to be called
                if assistant_message["tool_calls"]:
                    for tool_call in assistant_message["tool_calls"]:
                        tool_params = ToolExecutionParams(
                            tool_name=tool_call["function"]["name"],
                            tool_args=json.loads(tool_call["function"]["arguments"])
                        )

                        # Execute tool
                        tool_response = await workflow.execute_activity(
                            execute_tool,
                            args=[tool_params],
                            start_to_close_timeout=timedelta(minutes=1)
                        )

                        # Add tool response to messages
                        self.messages.append({
                            "role": "tool",
                            "tool_call_id": tool_call["id"],
                            "name": tool_call["function"]["name"],
                            "content": str(tool_response)
                        })

                    # Continue to let LLM process tool responses
                    continue

                # If no tools were called, break and wait for next user input
                break

    @workflow.signal
    async def user_input(self, input_data: Dict[str, Any]):
        """Signal handler for user input."""
        self.user_query = input_data.get("query", "")
        self.pptx_files = input_data.get("pptx_files", [])
        self.excel_files = input_data.get("excel_files", [])
        self.file_path_mapping = {
            os.path.basename(f): f 
            for f in self.pptx_files + self.excel_files
        }
        self.user_input_received = True

    @workflow.query
    def get_conversation_history(self) -> List[Dict[str, Any]]:
        """Query method to get the current conversation history."""
        return self.messages
