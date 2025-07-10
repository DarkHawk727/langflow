from unittest.mock import patch

import pytest
from composio import Action
from langflow.components.composio.sharepoint_composio import ComposioSharePointAPIComponent

from tests.base import DID_NOT_EXIST, ComponentTestBaseWithoutClient

from .test_base import MockComposioToolSet


class MockSharePointAction:
    SHARE_POINT_SHAREPOINT_CREATE_FOLDER = "SHARE_POINT_SHAREPOINT_CREATE_FOLDER"
    SHARE_POINT_SHAREPOINT_CREATE_LIST = "SHARE_POINT_SHAREPOINT_CREATE_LIST"
    SHARE_POINT_SHAREPOINT_CREATE_LIST_ITEM = "SHARE_POINT_SHAREPOINT_CREATE_LIST_ITEM"
    SHARE_POINT_SHAREPOINT_CREATE_USER = "SHARE_POINT_SHAREPOINT_CREATE_USER"
    SHARE_POINT_SHAREPOINT_FIND_USER = "SHARE_POINT_SHAREPOINT_FIND_USER"
    SHARE_POINT_SHAREPOINT_REMOVE_USER = "SHARE_POINT_SHAREPOINT_REMOVE_USER"


class TestSharePointComponent(ComponentTestBaseWithoutClient):
    @pytest.fixture(autouse=True)
    def mock_composio_toolset(self):
        with patch(
            "langflow.base.composio.composio_base.ComposioToolSet",
            MockComposioToolSet,
        ):
            yield

    @pytest.fixture
    def component_class(self):
        return ComposioSharePointAPIComponent

    @pytest.fixture
    def default_kwargs(self):
        return {
            "api_key": "",
            "entity_id": "default",
            "action": None,
        }

    @pytest.fixture
    def file_names_mapping(self):
        # Component not yet released, mark all versions as non-existent
        return [
            {"version": "1.0.0", "module": "composio", "file_name": DID_NOT_EXIST},
            {"version": "1.1.0", "module": "composio", "file_name": DID_NOT_EXIST},
            {"version": "1.2.0", "module": "composio", "file_name": DID_NOT_EXIST},
        ]

    def test_init(self, component_class, default_kwargs):
        component = component_class(**default_kwargs)
        assert component.display_name == "SharePoint"
        assert component.app_name == "sharepoint"
        # spot-check a couple of actions
        assert "SHARE_POINT_SHAREPOINT_CREATE_FOLDER" in component._actions_data
        assert "SHARE_POINT_SHAREPOINT_FIND_USER" in component._actions_data

    def test_execute_action_create_folder(self, component_class, default_kwargs, monkeypatch):
        # Mock Action enum
        monkeypatch.setattr(
            Action,
            "SHARE_POINT_SHAREPOINT_CREATE_FOLDER",
            MockSharePointAction.SHARE_POINT_SHAREPOINT_CREATE_FOLDER,
        )

        component = component_class(**default_kwargs)
        component.api_key = "test_key"
        component.action = [{"name": "Create Folder"}]
        component.document_library = "Shared Documents"
        component.folder_name = "TestFolder"
        component.relative_path = "folder/path"

        result = component.execute_action()
        assert result == {"result": "mocked response"}

    def test_execute_action_create_list(self, component_class, default_kwargs, monkeypatch):
        monkeypatch.setattr(
            Action,
            "SHARE_POINT_SHAREPOINT_CREATE_LIST",
            MockSharePointAction.SHARE_POINT_SHAREPOINT_CREATE_LIST,
        )

        component = component_class(**default_kwargs)
        component.api_key = "test_key"
        component.action = [{"name": "Create List"}]
        component.name = "TestList"
        component.template = "genericList"
        component.description = "A test list"

        result = component.execute_action()
        assert result == {"result": "mocked response"}

    def test_execute_action_create_list_item(self, component_class, default_kwargs, monkeypatch):
        monkeypatch.setattr(
            Action,
            "SHARE_POINT_SHAREPOINT_CREATE_LIST_ITEM",
            MockSharePointAction.SHARE_POINT_SHAREPOINT_CREATE_LIST_ITEM,
        )

        component = component_class(**default_kwargs)
        component.api_key = "test_key"
        component.action = [{"name": "Create List Item"}]
        component.list_name = "TestList"
        component.item_properties = {"Title": "Item1", "Value": 123}

        result = component.execute_action()
        assert result == {"result": "mocked response"}

    def test_execute_action_create_user(self, component_class, default_kwargs, monkeypatch):
        monkeypatch.setattr(
            Action,
            "SHARE_POINT_SHAREPOINT_CREATE_USER",
            MockSharePointAction.SHARE_POINT_SHAREPOINT_CREATE_USER,
        )

        component = component_class(**default_kwargs)
        component.api_key = "test_key"
        component.action = [{"name": "Create User"}]
        component.email = "user@example.com"
        component.login_name = "user_login"
        component.title = "Tester"

        result = component.execute_action()
        assert result == {"result": "mocked response"}

    def test_execute_action_find_user(self, component_class, default_kwargs, monkeypatch):
        monkeypatch.setattr(
            Action,
            "SHARE_POINT_SHAREPOINT_FIND_USER",
            MockSharePointAction.SHARE_POINT_SHAREPOINT_FIND_USER,
        )

        component = component_class(**default_kwargs)
        component.api_key = "test_key"
        component.action = [{"name": "Find User"}]
        component.email = "user@example.com"

        result = component.execute_action()
        assert result == {"result": "mocked response"}

    def test_execute_action_remove_user(self, component_class, default_kwargs, monkeypatch):
        monkeypatch.setattr(
            Action,
            "SHARE_POINT_SHAREPOINT_REMOVE_USER",
            MockSharePointAction.SHARE_POINT_SHAREPOINT_REMOVE_USER,
        )

        component = component_class(**default_kwargs)
        component.api_key = "test_key"
        component.action = [{"name": "Remove User"}]
        component.email = "user@example.com"

        result = component.execute_action()
        assert result == {"result": "mocked response"}

    def test_execute_action_invalid_action(self, component_class, default_kwargs):
        component = component_class(**default_kwargs)
        component.api_key = "test_key"
        component.action = [{"name": "Invalid Action"}]

        with pytest.raises(ValueError, match="Invalid action: Invalid Action"):
            component.execute_action()

    def test_update_build_config(self, component_class, default_kwargs):
        component = component_class(**default_kwargs)
        build_config = {
            "auth_link": {"value": "", "auth_tooltip": ""},
            "action": {
                "options": [],
                "helper_text": "",
                "helper_text_metadata": {},
            },
        }

        # No API key provided
        result = component.update_build_config(build_config, "", "api_key")
        assert result["auth_link"]["value"] == ""
        assert "Please provide a valid Composio API Key" in result["auth_link"]["auth_tooltip"]
        assert result["action"]["options"] == []

        # With API key
        component.api_key = "test_key"
        result = component.update_build_config(build_config, "test_key", "api_key")
        assert len(result["action"]["options"]) > 0
