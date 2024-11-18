from project import other_operation_prompt, select_month, validate_menu_option
from unittest.mock import patch
import pytest


def test_other_operation_prompt_yes():
    with patch('builtins.input', return_value='y'):
        assert other_operation_prompt() == None


def test_other_operation_prompt_no():
    with patch('builtins.input', return_value='n'):
        with pytest.raises(SystemExit) as excinfo: 
            other_operation_prompt()
        assert excinfo.type == SystemExit  
        assert str(excinfo.value) == "Exiting Expense Tracker...\n"


def test_select_month_valid():
    with patch('builtins.input', return_value='12'):
        assert select_month("") == (12, "December")


def test_validate_menu_option_valid():
    with patch('builtins.input', return_value='5'):
        assert validate_menu_option() == 5