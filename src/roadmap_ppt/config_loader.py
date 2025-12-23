"""
Configuration loader module.
Loads configuration from user's home directory or falls back to defaults.
"""

import os
import shutil
from pathlib import Path

# Import default config as fallback
from . import config as default_config


def get_config_dir():
    """Get the configuration directory path in user's home directory."""
    home = os.path.expanduser("~")
    config_dir = os.path.join(home, ".roadmap_ppt")
    return config_dir


def get_config_path():
    """Get the full path to the user's config file."""
    config_dir = get_config_dir()
    return os.path.join(config_dir, "config.py")


def create_default_config():
    """Create default config file in user's home directory if it doesn't exist."""
    config_path = get_config_path()
    config_dir = get_config_dir()
    
    # Create directory if it doesn't exist
    os.makedirs(config_dir, exist_ok=True)
    
    # If config file doesn't exist, copy default config
    if not os.path.exists(config_path):
        # Read default config content
        default_config_file = os.path.join(os.path.dirname(__file__), "config.py")
        if os.path.exists(default_config_file):
            shutil.copy2(default_config_file, config_path)
            print(f"Created default config file at: {config_path}")
            print("You can edit this file to customize your branding.")


def load_config():
    """
    Load configuration from user's home directory.
    Falls back to default config if user config doesn't exist.
    Creates default config file on first run.
    """
    config_path = get_config_path()
    
    # Create default config if it doesn't exist
    if not os.path.exists(config_path):
        create_default_config()
    
    # Try to load user config
    if os.path.exists(config_path):
        try:
            # Read and execute the config file
            with open(config_path, 'r', encoding='utf-8') as f:
                config_code = f.read()
            
            # Create a new module namespace
            import types
            user_config = types.ModuleType('user_config')
            
            # Execute config code in the module namespace
            exec(config_code, user_config.__dict__)
            
            # Return user config module
            return user_config
        except Exception as e:
            print(f"Warning: Could not load user config from {config_path}: {e}")
            print("Falling back to default configuration.")
            return default_config
    else:
        # Fall back to default config
        return default_config


# Load config module
config = load_config()

