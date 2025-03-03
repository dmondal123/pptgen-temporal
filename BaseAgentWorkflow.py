import _io
import json
from typing import Union, List, Dict, Any, Optional
from temporalio import workflow
from datetime import timedelta
import asyncio
import os
from dataclasses import dataclass

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

@dataclass
class UserInput:
    query: str
    pptx_files: List[str]
    excel_files: List[str]

@workflow.defn
class PPTAgentWorkflow:
    def __init__(self):
        self.messages: List[Dict[str, Any]] = []
        self.pptx_files: List[str] = []
        self.excel_files: List[str] = []
        self.file_path_mapping = {}
        self.memory: Dict[str, Any] = {}
        self.tools = define_tools()
        self.user_input_received = False
        self.user_query = ""

    @workflow.run
    async def run(self) -> List[Dict[str, Any]]:
        """Main workflow execution."""
        self.messages = [{
            "role": "system",
            "content": "You are an AI PowerPoint and Excel agent. You can view and modify PowerPoint slides and Excel sheets."
        }]
        return self.messages

    @workflow.query
    def get_conversation_history(self) -> List[Dict[str, Any]]:
        """Query method to get conversation history."""
        return self.messages

    @workflow.signal
    async def user_input(self, data: UserInput):
        """Signal method to receive user input."""
        self.pptx_files = data.pptx_files
        self.excel_files = data.excel_files
        
        # Add user message to conversation
        self.messages.append({
            "role": "user",
            "content": data.query
        })
        
        # Process the query and generate response
        # Add assistant response to messages
        response = await workflow.execute_activity(
            "call_llm",
            args=[{
                "messages": self.messages,
                "pptx_files": self.pptx_files,
                "excel_files": self.excel_files
            }],
            start_to_close_timeout=timedelta(minutes=5)
        )
        
        self.messages.append(response)

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
