from typing import Any

from composio import Action

from langflow.base.composio.composio_base import ComposioBaseComponent
from langflow.inputs.inputs import MessageTextInput, NestedDictInput
from langflow.logging import logger


class ComposioSharePointAPIComponent(ComposioBaseComponent):
    """SharePoint API component for interacting with SharePoint services."""

    display_name: str = "SharePoint"
    name = "SharePointAPI"
    icon = "SharePoint"
    documentation: str = "https://docs.composio.dev/tools/share-point"
    app_name = "sharepoint"

    # SharePoint-specific actions
    _actions_data = {
        "SHARE_POINT_SHAREPOINT_CREATE_FOLDER": {
            "display_name": "Create Folder",
            "action_fields": [
                "document_library",
                "folder_name",
                "relative_path",
            ],
        },
        "SHARE_POINT_SHAREPOINT_CREATE_LIST": {
            "display_name": "Create List",
            "action_fields": [
                "description",
                "name",
                "template",
            ],
        },
        "SHARE_POINT_SHAREPOINT_CREATE_LIST_ITEM": {
            "display_name": "Create List Item",
            "action_fields": [
                "item_properties",
                "list_name",
            ],
        },
        "SHARE_POINT_SHAREPOINT_CREATE_USER": {
            "display_name": "Create User",
            "action_fields": [
                "email",
                "login_name",
                "title",
            ],
        },
        "SHARE_POINT_SHAREPOINT_FIND_USER": {
            "display_name": "Find User",
            "action_fields": [
                "email",
            ],
        },
        "SHARE_POINT_SHAREPOINT_REMOVE_USER": {
            "display_name": "Remove User",
            "action_fields": [
                "email",
            ],
        },
    }

    _all_fields = {field for action_data in _actions_data.values() for field in action_data["action_fields"]}

    inputs = [
        *ComposioBaseComponent._base_inputs,
        # Folder creation
        MessageTextInput(
            name="document_library",
            display_name="Document Library",
            info="Name of the document library (default: 'Shared Documents')",
            show=False,
            required=False,
            advanced=False,
        ),
        MessageTextInput(
            name="folder_name",
            display_name="Folder Name",
            info="Name of the folder to create",
            show=False,
            required=True,
            advanced=False,
        ),
        MessageTextInput(
            name="relative_path",
            display_name="Relative Path",
            info="Relative path where the folder should be created",
            show=False,
            required=False,
            advanced=False,
        ),
        # List creation
        MessageTextInput(
            name="description",
            display_name="Description",
            info="Description for the new list",
            show=False,
            required=False,
            advanced=False,
        ),
        MessageTextInput(
            name="name",
            display_name="List Name",
            info="Name of the list to create",
            show=False,
            required=True,
            advanced=False,
        ),
        MessageTextInput(
            name="template",
            display_name="Template",
            info="Template type for the list",
            show=False,
            required=True,
            advanced=False,
        ),
        # List item creation
        NestedDictInput(  # Or DictInput, depending on your component base!
            name="item_properties",
            display_name="Item Properties",
            info="Properties for the item to be created",
            show=False,
            required=True,
            advanced=False,
        ),
        MessageTextInput(
            name="list_name",
            display_name="List Name",
            info="Name of the list to add the item to",
            show=False,
            required=True,
            advanced=False,
        ),
        # User management
        MessageTextInput(
            name="email",
            display_name="Email",
            info="Email address of the user",
            show=False,
            required=True,
            advanced=False,
        ),
        MessageTextInput(
            name="login_name",
            display_name="Login Name",
            info="Login name of the user",
            show=False,
            required=True,
            advanced=False,
        ),
        MessageTextInput(
            name="title",
            display_name="Title",
            info="Title of the user",
            show=False,
            required=True,
            advanced=False,
        ),
    ]

    def execute_action(self):
        """Execute SharePoint action and return response as Message."""
        toolset = self._build_wrapper()

        try:
            self._build_action_maps()
            display_name = self.action[0]["name"] if isinstance(self.action, list) and self.action else self.action
            action_key = self._display_to_key_map.get(display_name)
            if not action_key:
                msg = f"Invalid action: {display_name}"
                raise ValueError(msg)

            enum_name = getattr(Action, action_key)
            params = {}
            if action_key in self._actions_data:
                for field in self._actions_data[action_key]["action_fields"]:
                    value = getattr(self, field)
                    if value is None or value == "":
                        continue
                    params[field] = value

            result = toolset.execute_action(
                action=enum_name,
                params=params,
            )

            if not result.get("successful"):
                # Adjust this block if SharePoint error response looks different
                error_data = result.get("error", {})
                return {
                    "code": error_data.get("code"),
                    "message": error_data.get("message"),
                    "errors": error_data.get("errors", []),
                    "status": error_data.get("status"),
                }

            result_data = result.get("data", {})
            # If SharePoint ever adds a result_field/get_result_field logic, keep this,
            # but for now just return the full data.

        except Exception as e:
            logger.error(f"Error executing action: {e}")
            display_name = self.action[0]["name"] if isinstance(self.action, list) and self.action else str(self.action)
            msg = f"Failed to execute {display_name}: {e!s}"
            raise ValueError(msg) from e
        else:
            return result_data

    def update_build_config(self, build_config: dict, field_value: Any, field_name: str | None = None) -> dict:
        return super().update_build_config(build_config, field_value, field_name)

    def set_default_tools(self):
        self._default_tools = {
            "SHARE_POINT_SHAREPOINT_CREATE_USER",
            "SHARE_POINT_SHAREPOINT_FIND_USER",
        }
