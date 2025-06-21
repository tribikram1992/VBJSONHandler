"""
Driver for Windows UI automation using pywinauto.

Provides methods for interacting with Windows applications.
"""
import logging
import time
import os
import importlib
import sys
import traceback
from typing import Any, Optional, Dict, Union, Tuple, List
from pywinauto import Application, Desktop
from pywinauto.timings import wait_until, TimeoutError
import re

logger = logging.getLogger(__name__)

def get_framework_constants():
    """
    Get framework constants from the application's constants module.
    
    Returns:
        Dictionary of framework constants
    """
    constants = {}
    
    # First check if any framework_constants module is already imported
    for module_name in list(sys.modules.keys()):
        if module_name.endswith('.constants.framework_constants'):
            module = sys.modules[module_name]
            # Extract all uppercase constants
            constants = {name: getattr(module, name) for name in dir(module) 
                        if name.isupper() and not name.startswith('_')}
            
            # Also get the framework config
            if hasattr(module, 'FRAMEWORK_CONFIG_FILE'):
                try:
                    import yaml
                    with open(module.FRAMEWORK_CONFIG_FILE, 'r') as f:
                        framework_config = yaml.safe_load(f)
                        constants['FRAMEWORK_CONFIG'] = framework_config
                except Exception as e:
                    logger.warning(f"Could not load framework config: {e}")
            
            return constants
    
    # If not found, return empty dict
    return {}

class WindowsDriver:
    """
    Driver for Windows UI automation using pywinauto.
    
    Provides methods for interacting with Windows applications.
    Implements the singleton pattern to ensure only one instance exists.
    """
    
    # Class variable to track the single instance
    _instance = None
    _initialized = False
    
    def __new__(cls, config=None):
        """
        Implement singleton pattern to ensure only one driver instance exists.
        
        Args:
            config: Configuration manager (optional, only used for first initialization)
        
        Returns:
            WindowsDriver instance
        """
        if cls._instance is None:
            logger.info("Creating new WindowsDriver instance")
            cls._instance = super(WindowsDriver, cls).__new__(cls)
        return cls._instance
    
    def __init__(self, config=None):
        """
        Initialize the Windows driver.
        
        Args:
            config: Configuration manager (optional, only used for first initialization)
        """
        # Only initialize once
        if WindowsDriver._initialized:
            return
            
        if config is None:
            raise ValueError("Configuration manager must be provided for initial WindowsDriver initialization")
            
        self.config = config
        self.windows_config = config.get_all('windows_config')
        
        # Get framework constants
        constants = get_framework_constants()
        framework_config = constants.get('FRAMEWORK_CONFIG', {})
        
        # Get timeouts from framework config
        timeouts = framework_config.get('timeouts', {})
        self.default_timeout = timeouts.get('default', constants.get('DEFAULT_TIMEOUT', 30))
        self.element_timeout = timeouts.get('element', constants.get('ELEMENT_TIMEOUT', 20))
        self.window_timeout = timeouts.get('window', constants.get('WINDOW_TIMEOUT', 60))
        self.action_timeout = timeouts.get('action', constants.get('ACTION_TIMEOUT', 10))
        
        self.app = None
        self.desktop = Desktop(backend="uia")
        self.app_pid = None  # Store the application PID
        
        # Mark as initialized
        WindowsDriver._initialized = True
        
        logger.info("WindowsDriver initialized")
        logger.info(f"Timeouts: default={self.default_timeout}s, element={self.element_timeout}s, "
                   f"window={self.window_timeout}s, action={self.action_timeout}s")
    
    @classmethod
    def get_instance(cls, config=None):
        """
        Get the WindowsDriver instance.
        
        Args:
            config: Configuration manager (optional, only used for first initialization)
            
        Returns:
            WindowsDriver instance
        """
        if cls._instance is None:
            if config is None:
                raise ValueError("Configuration manager must be provided for initial WindowsDriver initialization")
            return cls(config)
        return cls._instance
    
    @classmethod
    def reset_instance(cls):
        """
        Reset the WindowsDriver instance.
        This is useful for testing or when you need to reinitialize the driver.
        """
        if cls._instance is not None and hasattr(cls._instance, 'app') and cls._instance.app is not None:
            try:
                cls._instance.app.kill()
            except Exception:
                pass
        cls._instance = None
        cls._initialized = False
        logger.info("WindowsDriver instance reset")
    
    def start_application(self, app_path: str = None, app_id: str = None, args: str = None, return_pid: bool = False) -> Union['WindowsDriver', Tuple[Any, int]]:
        """
        Start a Windows application.
        
        Args:
            app_path: Path to the application executable
            app_id: Application ID from configuration
            args: Command line arguments
            return_pid: Whether to return the process ID
            
        Returns:
            self: For method chaining, or (self, pid) if return_pid is True
        """
        # Check if app is already running
        if self.app:
            try:
                # Check if the application is still running
                if self.app.is_process_running():
                    logger.info("Application is already running, reusing existing instance")
                    if return_pid:
                        return self, self.app_pid
                    return self
            except Exception:
                # If checking fails, assume the app is not running
                pass
        
        if app_id and not app_path:
            app_path = self.config.get(f'windows_config.applications.{app_id}.path')
            if not app_path:
                raise ValueError(f"Application path not found for app_id: {app_id}")
        
        if not app_path:
            raise ValueError("Either app_path or app_id must be provided")
        
        logger.info(f"Starting Windows application: {app_path}")
        
        try:
            if args:
                self.app = Application(backend="uia").start(f"{app_path} {args}")
            else:
                self.app = Application(backend="uia").start(app_path)
            
            # Get and store the PID
            try:
                self.app_pid = self.app.process
                logger.info(f"Application started with PID: {self.app_pid}")
            except Exception as e:
                logger.warning(f"Could not get application PID: {e}")
                self.app_pid = None
            
            # Wait for the application to start
            time.sleep(2)  # Give the app a moment to initialize
            logger.info(f"Application started successfully")
            
            if return_pid:
                return self, self.app_pid
            return self
        except Exception as e:
            logger.error(f"Failed to start application: {str(e)}")
            logger.error(traceback.format_exc())
            if return_pid:
                return None, None
            raise
    
    def connect_to_application(self, process=None, handle=None, app_id=None) -> 'WindowsDriver':
        """
        Connect to a running Windows application.
        
        Args:
            process: Process ID or name
            handle: Window handle
            app_id: Application ID from configuration
            
        Returns:
            self: For method chaining
        """
        # Check if we're already connected
        if self.app:
            try:
                # Check if the application is still running
                if self.app.is_process_running():
                    logger.info("Already connected to application, reusing connection")
                    return self
            except Exception:
                # If checking fails, proceed with new connection
                pass
        
        if app_id and not process:
            process = self.config.get(f'windows_config.applications.{app_id}.process')
        
        try:
            if process:
                if isinstance(process, int):
                    logger.info(f"Connecting to application by PID: {process}")
                    self.app = Application(backend="uia").connect(process=process)
                    self.app_pid = process  # Store the PID
                else:
                    logger.info(f"Connecting to application by process name: {process}")
                    self.app = Application(backend="uia").connect(path=process)
                    # Try to get the PID
                    try:
                        self.app_pid = self.app.process
                        logger.info(f"Connected to application with PID: {self.app_pid}")
                    except Exception as e:
                        logger.warning(f"Could not get application PID: {e}")
                        self.app_pid = None
            elif handle:
                logger.info(f"Connecting to application by handle: {handle}")
                self.app = Application(backend="uia").connect(handle=handle)
                # Try to get the PID
                try:
                    self.app_pid = self.app.process
                    logger.info(f"Connected to application with PID: {self.app_pid}")
                except Exception as e:
                    logger.warning(f"Could not get application PID: {e}")
                    self.app_pid = None
            else:
                raise ValueError("Either process, handle, or app_id must be provided")
            
            logger.info("Connected to application successfully")
            return self
        except Exception as e:
            logger.error(f"Failed to connect to application: {str(e)}")
            logger.error(traceback.format_exc())
            raise
    
    def set_app_pid(self, pid: int) -> 'WindowsDriver':
        """
        Set the application PID manually.
        
        Args:
            pid: Process ID
            
        Returns:
            self: For method chaining
        """
        self.app_pid = pid
        logger.info(f"Application PID set to: {pid}")
        return self
    
    def get_app_pid(self) -> Optional[int]:
        """
        Get the application PID.
        
        Returns:
            int: Application PID or None if not available
        """
        return self.app_pid
    
    def find_window_by_pid(self, pid: int = None, title: str = None, class_name: str = None, timeout: int = None) -> Any:
        """
        Find a window by PID and optionally title and/or class name.
        
        Args:
            pid: Process ID (if None, will use self.app_pid)
            title: Window title (optional)
            class_name: Window class name (optional)
            timeout: Timeout in seconds
            
        Returns:
            Window object
        """
        timeout = timeout or self.window_timeout
        
        # Use the stored PID if none provided
        if pid is None:
            pid = self.app_pid
            
        if pid is None:
            raise ValueError("No PID provided and no application PID stored")
        
        logger.info(f"Finding window with PID: {pid}, title: {title}, class: {class_name}")
        
        try:
            # Use pywinauto Desktop to find all windows
            desktop = Desktop(backend="uia")
            
            start_time = time.time()
            while time.time() - start_time < timeout:
                # Get all windows
                windows = desktop.windows()
                
                # Filter by PID
                pid_windows = []
                for window in windows:
                    try:
                        if window.process_id() == pid:
                            pid_windows.append(window)
                    except Exception:
                        continue
                
                if not pid_windows:
                    logger.debug(f"No windows found for PID: {pid}")
                    time.sleep(1)
                    continue
                
                # If we have title or class_name, filter further
                if title or class_name:
                    for window in pid_windows:
                        try:
                            match = True
                            if title and title not in window.window_text():
                                match = False
                            if class_name and class_name != window.class_name():
                                match = False
                            if match:
                                logger.info(f"Found window: '{window.window_text()}' with PID: {pid}")
                                return window
                        except Exception as e:
                            logger.debug(f"Error checking window: {e}")
                    
                    # No match found, wait and try again
                    time.sleep(1)
                else:
                    # No title or class_name filter, return the first window
                    window = pid_windows[0]
                    logger.info(f"Found window: '{window.window_text()}' with PID: {pid}")
                    return window
            
            error_msg = f"Window not found after {timeout} seconds with PID: {pid}"
            if title:
                error_msg += f", title: {title}"
            if class_name:
                error_msg += f", class: {class_name}"
            logger.error(error_msg)
            raise TimeoutError(error_msg)
            
        except Exception as e:
            if isinstance(e, TimeoutError):
                raise
            logger.error(f"Error finding window by PID: {e}")
            logger.error(traceback.format_exc())
            raise
    
    def find_window(self, title=None, class_name=None, title_re=None, pid=None, timeout=None) -> Any:
        """
        Find a window by title and/or class name, with optional PID filtering.
        
        Args:
            title: Window title
            class_name: Window class name
            title_re: Window title regular expression
            pid: Process ID to filter by
            timeout: Timeout in seconds
            
        Returns:
            Window object
        """
        timeout = timeout or self.window_timeout
        
        # If PID is provided, use find_window_by_pid
        if pid is not None:
            return self.find_window_by_pid(pid=pid, title=title, class_name=class_name, timeout=timeout)
        
        # If app_pid is available and no specific criteria provided, try by PID first
        if self.app_pid and not title and not class_name and not title_re:
            try:
                return self.find_window_by_pid(pid=self.app_pid, timeout=timeout)
            except Exception as e:
                logger.debug(f"Could not find window by PID, falling back to other methods: {e}")
        
        if not title and not class_name and not title_re:
            raise ValueError("Either title, class_name, or title_re must be provided")
        
        criteria = {}
        if title:
            criteria['title'] = title
        if class_name:
            criteria['class_name'] = class_name
        if title_re:
            criteria['title_re'] = title_re
        
        logger.info(f"Finding window with criteria: {criteria}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                if not self.app:
                    # If no app connection, try to find the window using Desktop
                    desktop = Desktop(backend="uia")
                    windows = desktop.windows()
                    
                    for window in windows:
                        try:
                            match = True
                            if title and title != window.window_text():
                                match = False
                            if class_name and class_name != window.class_name():
                                match = False
                            if title_re and not re.match(title_re, window.window_text()):
                                match = False
                            
                            # If PID is available, check that too
                            if self.app_pid and match:
                                try:
                                    if window.process_id() != self.app_pid:
                                        match = False
                                except Exception:
                                    pass
                            
                            if match:
                                logger.info(f"Found window: '{window.window_text()}'")
                                return window
                        except Exception:
                            continue
                else:
                    # Use the app connection to find the window
                    window = self.app.window(**criteria)
                    if window.exists():
                        logger.info(f"Found window: '{window.window_text()}'")
                        return window
            except Exception as e:
                logger.debug(f"Window not found yet: {str(e)}")
            
            time.sleep(1)
        
        error_msg = f"Window not found after {timeout} seconds with criteria: {criteria}"
        logger.error(error_msg)
        raise TimeoutError(error_msg)
    
    def find_all_windows(self, pid=None) -> List[Any]:
        """
        Find all windows, optionally filtering by PID.
        
        Args:
            pid: Process ID to filter by (if None, will use self.app_pid if available)
            
        Returns:
            List of window objects
        """
        # Use the stored PID if none provided
        if pid is None and self.app_pid is not None:
            pid = self.app_pid
        
        try:
            # Use pywinauto Desktop to find all windows
            desktop = Desktop(backend="uia")
            windows = desktop.windows()
            
            # Filter by PID if provided
            if pid is not None:
                filtered_windows = []
                for window in windows:
                    try:
                        if window.process_id() == pid:
                            filtered_windows.append(window)
                    except Exception:
                        continue
                
                logger.info(f"Found {len(filtered_windows)} windows for PID: {pid}")
                return filtered_windows
            else:
                logger.info(f"Found {len(windows)} windows")
                return windows
                
        except Exception as e:
            logger.error(f"Error finding windows: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def find_element(self, parent_window, control_type=None, name=None, automation_id=None, class_name=None, timeout=None) -> Any:
        """
        Find an element in a window.
        
        Args:
            parent_window: Parent window object
            control_type: Control type
            name: Element name
            automation_id: Automation ID
            class_name: Class name
            timeout: Timeout in seconds
            
        Returns:
            Element object
        """
        timeout = timeout or self.element_timeout
        
        criteria = {}
        if control_type:
            criteria['control_type'] = control_type
        if name:
            criteria['title'] = name
        if automation_id:
            criteria['automation_id'] = automation_id
        if class_name:
            criteria['class_name'] = class_name
        
        if not criteria:
            raise ValueError("At least one search criterion must be provided")
        
        logger.info(f"Finding element with criteria: {criteria}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # Check if parent_window has child_window method
                if hasattr(parent_window, 'child_window'):
                    element = parent_window.child_window(**criteria)
                    if element.exists():
                        logger.info(f"Found element: {element.element_info.name}")
                        return element
                # If parent_window is a UIAWrapper, use descendants to find the element
                elif hasattr(parent_window, 'descendants'):
                    # Get all descendants
                    descendants = parent_window.descendants()
                    
                    # Filter by criteria
                    for element in descendants:
                        try:
                            match = True
                            # Check control_type
                            if control_type and hasattr(element, 'control_type'):
                                if callable(element.control_type):
                                    if element.control_type() != control_type:
                                        match = False
                                else:
                                    if element.control_type != control_type:
                                        match = False
                            
                            # Check name/title
                            if name and hasattr(element, 'window_text'):
                                if element.window_text() != name:
                                    match = False
                            
                            # Check automation_id
                            if automation_id and hasattr(element, 'automation_id'):
                                if callable(element.automation_id):
                                    if element.automation_id() != automation_id:
                                        match = False
                                else:
                                    if element.automation_id != automation_id:
                                        match = False
                            
                            # Check class_name
                            if class_name and hasattr(element, 'class_name'):
                                if callable(element.class_name):
                                    if element.class_name() != class_name:
                                        match = False
                                else:
                                    if element.class_name != class_name:
                                        match = False
                            
                            if match:
                                logger.info(f"Found element: {element.window_text() if hasattr(element, 'window_text') else 'Unknown'}")
                                return element
                        except Exception as e:
                            logger.debug(f"Error checking element: {e}")
                            continue
                else:
                    # If parent_window doesn't have either method, log an error
                    logger.error(f"Parent window doesn't have child_window or descendants method")
                    return None
            except Exception as e:
                logger.debug(f"Element not found yet: {str(e)}")
            
            time.sleep(0.5)
        
        error_msg = f"Element not found after {timeout} seconds with criteria: {criteria}"
        logger.error(error_msg)
        raise TimeoutError(error_msg)
    
    def find_dialog_window(self, title=None, class_name="#32770", timeout=5) -> Optional[Any]:
        """
        Find a dialog window that might be a popup message.
        
        Args:
            title: Dialog title or part of title
            class_name: Dialog class name (default is "#32770" for standard Windows dialogs)
            timeout: Timeout in seconds
            
        Returns:
            Dialog window object or None if not found
        """
        try:
            criteria = {}
            if title:
                criteria['title'] = title
            if class_name:
                criteria['class_name'] = class_name
                
            logger.info(f"Looking for dialog window with criteria: {criteria}")
            
            # Use Desktop to find all windows
            desktop = Desktop(backend="uia")
            
            start_time = time.time()
            while time.time() - start_time < timeout:
                windows = desktop.windows()
                
                for window in windows:
                    try:
                        match = True
                        if title and title not in window.window_text():
                            match = False
                        if class_name and class_name != window.class_name():
                            match = False
                            
                        if match:
                            logger.info(f"Found dialog window: '{window.window_text()}'")
                            return window
                    except Exception:
                        continue
                        
                time.sleep(0.5)
                
            logger.info(f"No dialog window found with criteria: {criteria}")
            return None
            
        except Exception as e:
            logger.error(f"Error finding dialog window: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def bring_window_to_foreground(self, window) -> 'WindowsDriver':
        """
        Bring a window to the foreground.
        
        Args:
            window: Window object
            
        Returns:
            self: For method chaining
        """
        try:
            if not window.is_active():
                logger.info(f"Bringing window to foreground: {window.window_text()}")
                window.set_focus()
                window.restore()
            return self
        except Exception as e:
            logger.error(f"Failed to bring window to foreground: {str(e)}")
            logger.error(traceback.format_exc())
            raise
    
    def take_screenshot(self, window, name) -> Optional[str]:
        """
        Take a screenshot of a window.
        
        Args:
            window: Window object
            name: Screenshot name
            
        Returns:
            Path to the screenshot
        """
        try:
            # Ensure window is in foreground before taking screenshot
            self.bring_window_to_foreground(window)
            
            # Import screenshot utility
            try:
                from src.main.utils.screenshot_utils import take_screenshot_windows
                return take_screenshot_windows(window, name)
            except ImportError:
                logger.warning("Screenshot utility not available, using fallback method")
                
                # Fallback to direct screenshot
                import os
                from datetime import datetime
                
                # Create screenshots directory if it doesn't exist
                screenshot_dir = os.path.join(os.getcwd(), 'reports', 'screenshots')
                if not os.path.exists(screenshot_dir):
                    os.makedirs(screenshot_dir)
                
                # Generate screenshot path with timestamp
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                screenshot_name = f"{name}_{timestamp}.png"
                screenshot_path = os.path.join(screenshot_dir, screenshot_name)
                
                # Take screenshot
                window.capture_as_image().save(screenshot_path)
                logger.info(f"Screenshot saved: {screenshot_path}")
                
                return screenshot_path
                
        except Exception as e:
            logger.error(f"Failed to take screenshot: {str(e)}")
            logger.error(traceback.format_exc())
            return None
    
    def close_application(self) -> 'WindowsDriver':
        """
        Close the application.
        
        Returns:
            self: For method chaining
        """
        if self.app:
            try:
                logger.info("Closing application")
                self.app.kill()
                self.app = None
                self.app_pid = None
            except Exception as e:
                logger.error(f"Failed to close application: {str(e)}")
                logger.error(traceback.format_exc())
        
        return self
    
    def click_element(self, element, coords=None) -> bool:
        """
        Click on an element.
        
        Args:
            element: Element to click
            coords: Optional coordinates to click (x, y)
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if coords:
                logger.info(f"Clicking element at coordinates: {coords}")
                element.click_input(coords=coords)
            else:
                logger.info("Clicking element")
                element.click_input()
            return True
        except Exception as e:
            logger.error(f"Failed to click element: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def set_text(self, element, text, clear_first=True) -> bool:
        """
        Set text in an input field.
        
        Args:
            element: Element to set text in
            text: Text to set
            clear_first: Whether to clear the field first
            
        Returns:
            True if successful, False otherwise
        """
        try:
            element.set_focus()
            
            if clear_first:
                logger.debug("Clearing text field")
                element.set_text("")
                
            logger.info(f"Setting text: {text}")
            element.type_keys(text, with_spaces=True)
            return True
        except Exception as e:
            logger.error(f"Failed to set text: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def get_text(self, element) -> Optional[str]:
        """
        Get text from an element.
        
        Args:
            element: Element to get text from
            
        Returns:
            Text if successful, None otherwise
        """
        try:
            text = element.window_text()
            logger.debug(f"Got text: {text}")
            return text
        except Exception as e:
            logger.error(f"Failed to get text: {str(e)}")
            logger.error(traceback.format_exc())
            return None
    
    def wait_for_element_visible(self, parent_window, control_type=None, name=None, automation_id=None, class_name=None, timeout=None) -> Any:
        """
        Wait for an element to become visible.
        
        Args:
            parent_window: Parent window object
            control_type: Control type
            name: Element name
            automation_id: Automation ID
            class_name: Class name
            timeout: Timeout in seconds
            
        Returns:
            Element object
        """
        timeout = timeout or self.element_timeout
        
        criteria = {}
        if control_type:
            criteria['control_type'] = control_type
        if name:
            criteria['title'] = name
        if automation_id:
            criteria['automation_id'] = automation_id
        if class_name:
            criteria['class_name'] = class_name
        
        if not criteria:
            raise ValueError("At least one search criterion must be provided")
        
        logger.info(f"Waiting for element to be visible with criteria: {criteria}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # Check if parent_window has child_window method
                if hasattr(parent_window, 'child_window'):
                    element = parent_window.child_window(**criteria)
                    if element.exists() and element.is_visible():
                        logger.info(f"Element is visible: {element.element_info.name}")
                        return element
                # If parent_window is a UIAWrapper, use descendants to find the element
                elif hasattr(parent_window, 'descendants'):
                    # Get all descendants
                    descendants = parent_window.descendants()
                    
                    # Filter by criteria
                    for element in descendants:
                        match = True
                        if control_type and element.control_type() != control_type:
                            match = False
                        if name and element.window_text() != name:
                            match = False
                        if automation_id and element.automation_id() != automation_id:
                            match = False
                        if class_name and element.class_name() != class_name:
                            match = False
                        
                        if match and element.is_visible():
                            logger.info(f"Element is visible: {element.window_text()}")
                            return element
                else:
                    # If parent_window doesn't have either method, log an error
                    logger.error(f"Parent window doesn't have child_window or descendants method")
                    return None
            except Exception as e:
                logger.debug(f"Element not visible yet: {str(e)}")
            
            time.sleep(0.5)
        
        error_msg = f"Element not visible after {timeout} seconds with criteria: {criteria}"
        logger.error(error_msg)
        raise TimeoutError(error_msg)
    
    def wait_for_element_enabled(self, parent_window, control_type=None, name=None, automation_id=None, class_name=None, timeout=None) -> Any:
        """
        Wait for an element to become enabled.
        
        Args:
            parent_window: Parent window object
            control_type: Control type
            name: Element name
            automation_id: Automation ID
            class_name: Class name
            timeout: Timeout in seconds
            
        Returns:
            Element object
        """
        timeout = timeout or self.element_timeout
        
        criteria = {}
        if control_type:
            criteria['control_type'] = control_type
        if name:
            criteria['title'] = name
        if automation_id:
            criteria['automation_id'] = automation_id
        if class_name:
            criteria['class_name'] = class_name
        
        if not criteria:
            raise ValueError("At least one search criterion must be provided")
        
        logger.info(f"Waiting for element to be enabled with criteria: {criteria}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # Check if parent_window has child_window method
                if hasattr(parent_window, 'child_window'):
                    element = parent_window.child_window(**criteria)
                    if element.exists() and element.is_enabled():
                        logger.info(f"Element is enabled: {element.element_info.name}")
                        return element
                # If parent_window is a UIAWrapper, use descendants to find the element
                elif hasattr(parent_window, 'descendants'):
                    # Get all descendants
                    descendants = parent_window.descendants()
                    
                    # Filter by criteria
                    for element in descendants:
                        match = True
                        if control_type and element.control_type() != control_type:
                            match = False
                        if name and element.window_text() != name:
                            match = False
                        if automation_id and element.automation_id() != automation_id:
                            match = False
                        if class_name and element.class_name() != class_name:
                            match = False
                        
                        if match and element.is_enabled():
                            logger.info(f"Element is enabled: {element.window_text()}")
                            return element
                else:
                    # If parent_window doesn't have either method, log an error
                    logger.error(f"Parent window doesn't have child_window or descendants method")
                    return None
            except Exception as e:
                logger.debug(f"Element not enabled yet: {str(e)}")
            
            time.sleep(0.5)
        
        error_msg = f"Element not enabled after {timeout} seconds with criteria: {criteria}"
        logger.error(error_msg)
        raise TimeoutError(error_msg)
    
    def get_all_child_windows(self, parent_window) -> List[Any]:
        """
        Get all child windows of a parent window.
        
        Args:
            parent_window: Parent window object
            
        Returns:
            List of child window objects
        """
        try:
            if hasattr(parent_window, 'children'):
                children = parent_window.children()
                logger.info(f"Found {len(children)} child windows")
                return children
            else:
                logger.warning("Parent window doesn't have children method")
                return []
        except Exception as e:
            logger.error(f"Error getting child windows: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_window_text(self, window) -> str:
        """
        Get the text of a window.
        
        Args:
            window: Window object
            
        Returns:
            Window text
        """
        try:
            text = window.window_text()
            logger.debug(f"Window text: {text}")
            return text
        except Exception as e:
            logger.error(f"Error getting window text: {e}")
            logger.error(traceback.format_exc())
            return ""
    
    def extract_text_from_dialog(self, dialog_window) -> str:
        """
        Extract text from a dialog window, including all static text controls.
        
        Args:
            dialog_window: Dialog window object
            
        Returns:
            Combined text from all static text controls
        """
        try:
            # First try to get text from Static controls
            text_elements = []
            
            # Try different control types that might contain text
            for control_type in ["Text", "Static", "Label", "Edit"]:
                try:
                    controls = dialog_window.children(control_type=control_type)
                    for control in controls:
                        if hasattr(control, 'window_text'):
                            text = control.window_text()
                            if text and text != dialog_window.window_text():
                                text_elements.append(text)
                except Exception:
                    pass
            
            # If no text found in specific controls, try all descendants
            if not text_elements and hasattr(dialog_window, 'descendants'):
                for element in dialog_window.descendants():
                    if hasattr(element, 'window_text'):
                        text = element.window_text()
                        if text and text != dialog_window.window_text() and text not in ["OK", "Cancel", "Yes", "No"]:
                            text_elements.append(text)
            
            # If still no text found, use the window title
            if not text_elements and hasattr(dialog_window, 'window_text'):
                window_title = dialog_window.window_text()
                if window_title:
                    text_elements.append(window_title)
            
            # Combine all text elements
            combined_text = " ".join(text_elements)
            logger.info(f"Extracted text from dialog: {combined_text}")
            return combined_text
            
        except Exception as e:
            logger.error(f"Error extracting text from dialog: {e}")
            logger.error(traceback.format_exc())
            return ""
    
    def maximize_window(self, window) -> bool:
        """
        Maximize a window.
        
        Args:
            window: Window object
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info("Maximizing window")
            window.maximize()
            return True
        except Exception as e:
            logger.error(f"Failed to maximize window: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def minimize_window(self, window) -> bool:
        """
        Minimize a window.
        
        Args:
            window: Window object
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info("Minimizing window")
            window.minimize()
            return True
        except Exception as e:
            logger.error(f"Failed to minimize window: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def restore_window(self, window) -> bool:
        """
        Restore a window.
        
        Args:
            window: Window object
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info("Restoring window")
            window.restore()
            return True
        except Exception as e:
            logger.error(f"Failed to restore window: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_maximized(self, window) -> bool:
        """
        Check if a window is maximized.
        
        Args:
            window: Window object
            
        Returns:
            True if maximized, False otherwise
        """
        try:
            if hasattr(window, 'is_maximized'):
                return window.is_maximized()
            else:
                logger.warning("Window doesn't have is_maximized method")
                return False
        except Exception as e:
            logger.error(f"Error checking if window is maximized: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_minimized(self, window) -> bool:
        """
        Check if a window is minimized.
        
        Args:
            window: Window object
            
        Returns:
            True if minimized, False otherwise
        """
        try:
            if hasattr(window, 'is_minimized'):
                return window.is_minimized()
            else:
                logger.warning("Window doesn't have is_minimized method")
                return False
        except Exception as e:
            logger.error(f"Error checking if window is minimized: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_visible(self, window) -> bool:
        """
        Check if a window is visible.
        
        Args:
            window: Window object
            
        Returns:
            True if visible, False otherwise
        """
        try:
            if hasattr(window, 'is_visible'):
                return window.is_visible()
            else:
                logger.warning("Window doesn't have is_visible method")
                return False
        except Exception as e:
            logger.error(f"Error checking if window is visible: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_enabled(self, window) -> bool:
        """
        Check if a window is enabled.
        
        Args:
            window: Window object
            
        Returns:
            True if enabled, False otherwise
        """
        try:
            if hasattr(window, 'is_enabled'):
                return window.is_enabled()
            else:
                logger.warning("Window doesn't have is_enabled method")
                return False
        except Exception as e:
            logger.error(f"Error checking if window is enabled: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def send_keystrokes(self, window, keystrokes) -> bool:
        """
        Send keystrokes to a window.
        
        Args:
            window: Window object
            keystrokes: Keystrokes to send
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Sending keystrokes: {keystrokes}")
            window.type_keys(keystrokes)
            return True
        except Exception as e:
            logger.error(f"Failed to send keystrokes: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def right_click_element(self, element, coords=None) -> bool:
        """
        Right-click on an element.
        
        Args:
            element: Element to right-click
            coords: Optional coordinates to right-click (x, y)
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if coords:
                logger.info(f"Right-clicking element at coordinates: {coords}")
                element.right_click_input(coords=coords)
            else:
                logger.info("Right-clicking element")
                element.right_click_input()
            return True
        except Exception as e:
            logger.error(f"Failed to right-click element: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def double_click_element(self, element, coords=None) -> bool:
        """
        Double-click on an element.
        
        Args:
            element: Element to double-click
            coords: Optional coordinates to double-click (x, y)
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if coords:
                logger.info(f"Double-clicking element at coordinates: {coords}")
                element.double_click_input(coords=coords)
            else:
                logger.info("Double-clicking element")
                element.double_click_input()
            return True
        except Exception as e:
            logger.error(f"Failed to double-click element: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    def select_combobox_item(self, combobox, item_text) -> bool:
        """
        Select an item from a combobox by text.
        
        Args:
            combobox: Combobox element
            item_text: Text of the item to select
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Selecting item '{item_text}' from combobox")
            
            # First click to expand the combobox
            combobox.click_input()
            time.sleep(0.5)
            
            # Try to find and click the item
            try:
                # Method 1: Try to find the item as a direct child of the combobox
                item = combobox.child_window(title=item_text)
                if item.exists():
                    item.click_input()
                    logger.info(f"Selected item '{item_text}' (method 1)")
                    return True
            except Exception as e:
                logger.debug(f"Method 1 failed: {e}")
            
            try:
                # Method 2: Look for a listbox or list item with the text
                # Find any listbox that might be the dropdown
                listboxes = self.desktop.windows(control_type="ListBox")
                for listbox in listboxes:
                    try:
                        item = listbox.child_window(title=item_text)
                        if item.exists():
                            item.click_input()
                            logger.info(f"Selected item '{item_text}' (method 2)")
                            return True
                    except Exception:
                        continue
            except Exception as e:
                logger.debug(f"Method 2 failed: {e}")
            
            try:
                # Method 3: Try to use keyboard navigation
                from pywinauto.keyboard import send_keys
                
                # Type the first few characters of the item text to select it
                send_keys(item_text[:3])
                time.sleep(0.5)
                send_keys("{ENTER}")
                logger.info(f"Tried to select item '{item_text}' using keyboard (method 3)")
                return True
            except Exception as e:
                logger.debug(f"Method 3 failed: {e}")
            
            logger.warning(f"Could not find or select item '{item_text}' in combobox")
            return False
            
        except Exception as e:
            logger.error(f"Failed to select combobox item: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def check_element_exists(self, parent_window, control_type=None, name=None, automation_id=None, class_name=None, timeout=1) -> bool:
        """
        Check if an element exists without waiting for the full timeout.
        
        Args:
            parent_window: Parent window object
            control_type: Control type
            name: Element name
            automation_id: Automation ID
            class_name: Class name
            timeout: Short timeout in seconds (default is 1)
            
        Returns:
            True if element exists, False otherwise
        """
        try:
            criteria = {}
            if control_type:
                criteria['control_type'] = control_type
            if name:
                criteria['title'] = name
            if automation_id:
                criteria['automation_id'] = automation_id
            if class_name:
                criteria['class_name'] = class_name
            
            if not criteria:
                raise ValueError("At least one search criterion must be provided")
            
            # Check if parent_window has child_window method
            if hasattr(parent_window, 'child_window'):
                element = parent_window.child_window(**criteria)
                return element.exists()
            # If parent_window is a UIAWrapper, use descendants to find the element
            elif hasattr(parent_window, 'descendants'):
                # Get all descendants
                descendants = parent_window.descendants()
                
                # Filter by criteria
                for element in descendants:
                    try:
                        match = True
                        if control_type and element.control_type() != control_type:
                            match = False
                        if name and element.window_text() != name:
                            match = False
                        if automation_id and element.automation_id() != automation_id:
                            match = False
                        if class_name and element.class_name() != class_name:
                            match = False
                        
                        if match:
                            return True
                    except Exception:
                        continue
                
                return False
            else:
                # If parent_window doesn't have either method, log an error
                logger.error(f"Parent window doesn't have child_window or descendants method")
                return False
                
        except Exception as e:
            logger.debug(f"Error checking if element exists: {e}")
            return False
    
    def wait_for_element_to_disappear(self, parent_window, control_type=None, name=None, automation_id=None, class_name=None, timeout=None) -> bool:
        """
        Wait for an element to disappear.
        
        Args:
            parent_window: Parent window object
            control_type: Control type
            name: Element name
            automation_id: Automation ID
            class_name: Class name
            timeout: Timeout in seconds
            
        Returns:
            True if element disappears, False if it's still visible after timeout
        """
        timeout = timeout or self.element_timeout
        
        criteria = {}
        if control_type:
            criteria['control_type'] = control_type
        if name:
            criteria['title'] = name
        if automation_id:
            criteria['automation_id'] = automation_id
        if class_name:
            criteria['class_name'] = class_name
        
        if not criteria:
            raise ValueError("At least one search criterion must be provided")
        
        logger.info(f"Waiting for element to disappear with criteria: {criteria}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            if not self.check_element_exists(parent_window, control_type, name, automation_id, class_name, 1):
                logger.info("Element has disappeared")
                return True
            time.sleep(0.5)
        
        logger.warning(f"Element still visible after {timeout} seconds")
        return False
    
    def wait_for_window_to_close(self, window, timeout=None) -> bool:
        """
        Wait for a window to close.
        
        Args:
            window: Window object
            timeout: Timeout in seconds
            
        Returns:
            True if window closes, False if it's still open after timeout
        """
        timeout = timeout or self.window_timeout
        
        logger.info("Waiting for window to close")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                if not window.exists():
                    logger.info("Window has closed")
                    return True
            except Exception:
                # If checking fails, assume the window is closed
                logger.info("Window has closed (exception while checking)")
                return True
            time.sleep(0.5)
        
        logger.warning(f"Window still open after {timeout} seconds")
        return False
    
    def wait_for_process_to_exit(self, pid=None, timeout=None) -> bool:
        """
        Wait for a process to exit.
        
        Args:
            pid: Process ID (if None, will use self.app_pid)
            timeout: Timeout in seconds
            
        Returns:
            True if process exits, False if it's still running after timeout
        """
        timeout = timeout or self.window_timeout
        
        # Use the stored PID if none provided
        if pid is None:
            pid = self.app_pid
            
        if pid is None:
            raise ValueError("No PID provided and no application PID stored")
        
        logger.info(f"Waiting for process {pid} to exit")
        
        try:
            import psutil
            
            start_time = time.time()
            while time.time() - start_time < timeout:
                if not psutil.pid_exists(pid):
                    logger.info(f"Process {pid} has exited")
                    return True
                time.sleep(0.5)
            
            logger.warning(f"Process {pid} still running after {timeout} seconds")
            return False
            
        except ImportError:
            logger.warning("psutil not available, using alternative method")
            
            # Alternative method using pywinauto
            if self.app:
                start_time = time.time()
                while time.time() - start_time < timeout:
                    try:
                        if not self.app.is_process_running():
                            logger.info(f"Process {pid} has exited")
                            return True
                    except Exception:
                        # If checking fails, assume the process is not running
                        logger.info(f"Process {pid} has exited (exception while checking)")
                        return True
                    time.sleep(0.5)
                
                logger.warning(f"Process {pid} still running after {timeout} seconds")
                return False
            else:
                logger.warning("No app connection available to check process status")
                return False
        except Exception as e:
            logger.error(f"Error waiting for process to exit: {e}")
            logger.error(traceback.format_exc())
            return False
    def get_all_windows_by_title(self, title_pattern, regex=False) -> List[Any]:
        """
        Get all windows matching a title pattern.
        
        Args:
            title_pattern: Title pattern to match
            regex: Whether to use regex matching
            
        Returns:
            List of matching window objects
        """
        try:
            # Use Desktop to find all windows
            desktop = Desktop(backend="uia")
            windows = desktop.windows()
            
            matching_windows = []
            for window in windows:
                try:
                    window_title = window.window_text()
                    if regex:
                        if re.search(title_pattern, window_title):
                            matching_windows.append(window)
                    else:
                        if title_pattern in window_title:
                            matching_windows.append(window)
                except Exception:
                    continue
            
            logger.info(f"Found {len(matching_windows)} windows matching pattern: {title_pattern}")
            return matching_windows
            
        except Exception as e:
            logger.error(f"Error finding windows by title: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_all_windows_by_class(self, class_name) -> List[Any]:
        """
        Get all windows with a specific class name.
        
        Args:
            class_name: Class name to match
            
        Returns:
            List of matching window objects
        """
        try:
            # Use Desktop to find all windows
            desktop = Desktop(backend="uia")
            windows = desktop.windows()
            
            matching_windows = []
            for window in windows:
                try:
                    if window.class_name() == class_name:
                        matching_windows.append(window)
                except Exception:
                    continue
            
            logger.info(f"Found {len(matching_windows)} windows with class name: {class_name}")
            return matching_windows
            
        except Exception as e:
            logger.error(f"Error finding windows by class: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_active_window(self) -> Optional[Any]:
        """
        Get the currently active window.
        
        Returns:
            Active window object or None if not found
        """
        try:
            # Use Desktop to find all windows
            desktop = Desktop(backend="uia")
            windows = desktop.windows()
            
            for window in windows:
                try:
                    if window.is_active():
                        logger.info(f"Found active window: '{window.window_text()}'")
                        return window
                except Exception:
                    continue
            
            logger.warning("No active window found")
            return None
            
        except Exception as e:
            logger.error(f"Error finding active window: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def wait_for_window_title(self, title_pattern, regex=False, timeout=None) -> Optional[Any]:
        """
        Wait for a window with a specific title to appear.
        
        Args:
            title_pattern: Title pattern to match
            regex: Whether to use regex matching
            timeout: Timeout in seconds
            
        Returns:
            Window object or None if not found
        """
        timeout = timeout or self.window_timeout
        
        logger.info(f"Waiting for window with title pattern: {title_pattern}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            windows = self.get_all_windows_by_title(title_pattern, regex)
            if windows:
                logger.info(f"Found window with title pattern: {title_pattern}")
                return windows[0]
            time.sleep(0.5)
        
        logger.warning(f"Window with title pattern '{title_pattern}' not found after {timeout} seconds")
        return None
    
    def wait_for_window_class(self, class_name, timeout=None) -> Optional[Any]:
        """
        Wait for a window with a specific class name to appear.
        
        Args:
            class_name: Class name to match
            timeout: Timeout in seconds
            
        Returns:
            Window object or None if not found
        """
        timeout = timeout or self.window_timeout
        
        logger.info(f"Waiting for window with class name: {class_name}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            windows = self.get_all_windows_by_class(class_name)
            if windows:
                logger.info(f"Found window with class name: {class_name}")
                return windows[0]
            time.sleep(0.5)
        
        logger.warning(f"Window with class name '{class_name}' not found after {timeout} seconds")
        return None
    
    def is_element_visible(self, element) -> bool:
        """
        Check if an element is visible.
        
        Args:
            element: Element to check
            
        Returns:
            True if visible, False otherwise
        """
        try:
            return element.is_visible()
        except Exception as e:
            logger.debug(f"Error checking if element is visible: {e}")
            return False
    
    def is_element_enabled(self, element) -> bool:
        """
        Check if an element is enabled.
        
        Args:
            element: Element to check
            
        Returns:
            True if enabled, False otherwise
        """
        try:
            return element.is_enabled()
        except Exception as e:
            logger.debug(f"Error checking if element is enabled: {e}")
            return False
    
    def get_element_rectangle(self, element) -> Optional[Any]:
        """
        Get the rectangle (position and size) of an element.
        
        Args:
            element: Element to get rectangle for
            
        Returns:
            Rectangle object or None if failed
        """
        try:
            return element.rectangle()
        except Exception as e:
            logger.error(f"Error getting element rectangle: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_element_center(self, element) -> Optional[Tuple[int, int]]:
        """
        Get the center coordinates of an element.
        
        Args:
            element: Element to get center for
            
        Returns:
            Tuple of (x, y) coordinates or None if failed
        """
        try:
            rect = element.rectangle()
            center_x = (rect.left + rect.right) // 2
            center_y = (rect.top + rect.bottom) // 2
            return (center_x, center_y)
        except Exception as e:
            logger.error(f"Error getting element center: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def drag_and_drop(self, source_element, target_element) -> bool:
        """
        Perform drag and drop from source element to target element.
        
        Args:
            source_element: Source element
            target_element: Target element
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info("Performing drag and drop")
            source_element.drag_mouse_input(target_element)
            return True
        except Exception as e:
            logger.error(f"Failed to perform drag and drop: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def drag_and_drop_by_offset(self, element, x_offset, y_offset) -> bool:
        """
        Perform drag and drop from element by offset.
        
        Args:
            element: Element to drag
            x_offset: X offset
            y_offset: Y offset
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Performing drag and drop by offset: ({x_offset}, {y_offset})")
            
            # Get element center
            center = self.get_element_center(element)
            if not center:
                logger.error("Could not get element center")
                return False
            
            # Calculate target coordinates
            target_x = center[0] + x_offset
            target_y = center[1] + y_offset
            
            # Perform drag and drop
            element.press_mouse_input()
            time.sleep(0.1)
            element.move_mouse_input(coords=(target_x, target_y))
            time.sleep(0.1)
            element.release_mouse_input(coords=(target_x, target_y))
            
            return True
        except Exception as e:
            logger.error(f"Failed to perform drag and drop by offset: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def hover_element(self, element) -> bool:
        """
        Hover over an element.
        
        Args:
            element: Element to hover over
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info("Hovering over element")
            element.move_mouse_input()
            return True
        except Exception as e:
            logger.error(f"Failed to hover over element: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def scroll_element(self, element, direction="down", amount=10) -> bool:
        """
        Scroll an element.
        
        Args:
            element: Element to scroll
            direction: Scroll direction ("up", "down", "left", "right")
            amount: Scroll amount
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Scrolling element {direction} by {amount}")
            
            if direction == "down":
                element.scroll(amount_of_scrolling=amount, direction="down")
            elif direction == "up":
                element.scroll(amount_of_scrolling=amount, direction="up")
            elif direction == "left":
                element.scroll(amount_of_scrolling=amount, direction="left")
            elif direction == "right":
                element.scroll(amount_of_scrolling=amount, direction="right")
            else:
                logger.error(f"Invalid scroll direction: {direction}")
                return False
            
            return True
        except Exception as e:
            logger.error(f"Failed to scroll element: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    def get_element_children(self, element) -> List[Any]:
        """
        Get all children of an element.
        
        Args:
            element: Parent element
            
        Returns:
            List of child elements
        """
        try:
            if hasattr(element, 'children'):
                children = element.children()
                logger.info(f"Found {len(children)} children")
                return children
            else:
                logger.warning("Element doesn't have children method")
                return []
        except Exception as e:
            logger.error(f"Error getting element children: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_element_descendants(self, element) -> List[Any]:
        """
        Get all descendants of an element.
        
        Args:
            element: Parent element
            
        Returns:
            List of descendant elements
        """
        try:
            if hasattr(element, 'descendants'):
                descendants = element.descendants()
                logger.info(f"Found {len(descendants)} descendants")
                return descendants
            else:
                logger.warning("Element doesn't have descendants method")
                return []
        except Exception as e:
            logger.error(f"Error getting element descendants: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_element_parent(self, element) -> Optional[Any]:
        """
        Get the parent of an element.
        
        Args:
            element: Child element
            
        Returns:
            Parent element or None if not found
        """
        try:
            if hasattr(element, 'parent'):
                parent = element.parent()
                logger.info(f"Found parent: {parent.window_text() if hasattr(parent, 'window_text') else 'Unknown'}")
                return parent
            else:
                logger.warning("Element doesn't have parent method")
                return None
        except Exception as e:
            logger.error(f"Error getting element parent: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_element_siblings(self, element) -> List[Any]:
        """
        Get all siblings of an element.
        
        Args:
            element: Element to get siblings for
            
        Returns:
            List of sibling elements
        """
        try:
            parent = self.get_element_parent(element)
            if parent:
                siblings = self.get_element_children(parent)
                # Remove the original element from the list
                siblings = [sibling for sibling in siblings if sibling != element]
                logger.info(f"Found {len(siblings)} siblings")
                return siblings
            else:
                logger.warning("Could not find parent element")
                return []
        except Exception as e:
            logger.error(f"Error getting element siblings: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_element_by_path(self, start_element, path) -> Optional[Any]:
        """
        Get an element by path from a starting element.
        
        Args:
            start_element: Starting element
            path: Path to the element (list of indices or criteria dictionaries)
            
        Returns:
            Element or None if not found
        """
        try:
            current_element = start_element
            
            for step in path:
                if isinstance(step, int):
                    # If step is an integer, get the child at that index
                    children = self.get_element_children(current_element)
                    if 0 <= step < len(children):
                        current_element = children[step]
                    else:
                        logger.error(f"Index {step} out of range")
                        return None
                elif isinstance(step, dict):
                    # If step is a dictionary, use it as criteria for find_element
                    control_type = step.get('control_type')
                    name = step.get('name')
                    automation_id = step.get('automation_id')
                    class_name = step.get('class_name')
                    
                    current_element = self.find_element(
                        current_element,
                        control_type=control_type,
                        name=name,
                        automation_id=automation_id,
                        class_name=class_name
                    )
                    
                    if not current_element:
                        logger.error(f"Element not found with criteria: {step}")
                        return None
                else:
                    logger.error(f"Invalid path step: {step}")
                    return None
            
            return current_element
        except Exception as e:
            logger.error(f"Error getting element by path: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_element_attribute(self, element, attribute_name) -> Optional[Any]:
        """
        Get an attribute of an element.
        
        Args:
            element: Element to get attribute from
            attribute_name: Name of the attribute
            
        Returns:
            Attribute value or None if not found
        """
        try:
            if hasattr(element, attribute_name):
                attr = getattr(element, attribute_name)
                if callable(attr):
                    return attr()
                else:
                    return attr
            else:
                logger.warning(f"Element doesn't have attribute: {attribute_name}")
                return None
        except Exception as e:
            logger.error(f"Error getting element attribute: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def wait_for_element_attribute(self, element, attribute_name, expected_value, timeout=None) -> bool:
        """
        Wait for an element attribute to have a specific value.
        
        Args:
            element: Element to check
            attribute_name: Name of the attribute
            expected_value: Expected value of the attribute
            timeout: Timeout in seconds
            
        Returns:
            True if attribute has expected value, False otherwise
        """
        timeout = timeout or self.element_timeout
        
        logger.info(f"Waiting for element attribute {attribute_name} to be {expected_value}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            attr_value = self.get_element_attribute(element, attribute_name)
            if attr_value == expected_value:
                logger.info(f"Element attribute {attribute_name} is now {expected_value}")
                return True
            time.sleep(0.5)
        
        logger.warning(f"Element attribute {attribute_name} is still not {expected_value} after {timeout} seconds")
        return False
    
    def wait_for_element_attribute_contains(self, element, attribute_name, expected_substring, timeout=None) -> bool:
        """
        Wait for an element attribute to contain a specific substring.
        
        Args:
            element: Element to check
            attribute_name: Name of the attribute
            expected_substring: Expected substring in the attribute value
            timeout: Timeout in seconds
            
        Returns:
            True if attribute contains expected substring, False otherwise
        """
        timeout = timeout or self.element_timeout
        
        logger.info(f"Waiting for element attribute {attribute_name} to contain '{expected_substring}'")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            attr_value = self.get_element_attribute(element, attribute_name)
            if attr_value and expected_substring in str(attr_value):
                logger.info(f"Element attribute {attribute_name} now contains '{expected_substring}'")
                return True
            time.sleep(0.5)
        
        logger.warning(f"Element attribute {attribute_name} still doesn't contain '{expected_substring}' after {timeout} seconds")
        return False
    
    def get_element_properties(self, element) -> Dict[str, Any]:
        """
        Get all properties of an element.
        
        Args:
            element: Element to get properties from
            
        Returns:
            Dictionary of property names and values
        """
        try:
            properties = {}
            
            # Common properties to check
            common_properties = [
                'window_text', 'class_name', 'control_type', 'automation_id',
                'is_visible', 'is_enabled', 'rectangle', 'process_id'
            ]
            
            for prop in common_properties:
                if hasattr(element, prop):
                    try:
                        attr = getattr(element, prop)
                        if callable(attr):
                            properties[prop] = attr()
                        else:
                            properties[prop] = attr
                    except Exception as e:
                        logger.debug(f"Error getting property {prop}: {e}")
            
            logger.info(f"Got {len(properties)} properties for element")
            return properties
        except Exception as e:
            logger.error(f"Error getting element properties: {e}")
            logger.error(traceback.format_exc())
            return {}
    
    def highlight_element(self, element, duration=1.0) -> bool:
        """
        Highlight an element by drawing a rectangle around it.
        
        Args:
            element: Element to highlight
            duration: Duration in seconds to show the highlight
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Highlighting element for {duration} seconds")
            element.draw_outline(colour='red', thickness=2)
            time.sleep(duration)
            return True
        except Exception as e:
            logger.error(f"Failed to highlight element: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def get_element_path(self, element) -> List[Dict[str, Any]]:
        """
        Get the path to an element from the root.
        
        Args:
            element: Element to get path for
            
        Returns:
            List of dictionaries with element information, from root to the element
        """
        try:
            path = []
            current = element
            
            while current:
                # Get element properties
                props = self.get_element_properties(current)
                path.insert(0, props)
                
                # Move to parent
                current = self.get_element_parent(current)
            
            logger.info(f"Got element path with {len(path)} nodes")
            return path
        except Exception as e:
            logger.error(f"Error getting element path: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_element_by_coordinates(self, window, x, y) -> Optional[Any]:
        """
        Get the element at specific coordinates.
        
        Args:
            window: Window to search in
            x: X coordinate
            y: Y coordinate
            
        Returns:
            Element at the coordinates or None if not found
        """
        try:
            logger.info(f"Getting element at coordinates: ({x}, {y})")
            element = window.from_point(x, y)
            if element:
                logger.info(f"Found element: {element.window_text() if hasattr(element, 'window_text') else 'Unknown'}")
            else:
                logger.warning("No element found at coordinates")
            return element
        except Exception as e:
            logger.error(f"Error getting element by coordinates: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_screen_resolution(self) -> Tuple[int, int]:
        """
        Get the screen resolution.
        
        Returns:
            Tuple of (width, height)
        """
        try:
            import ctypes
            user32 = ctypes.windll.user32
            width = user32.GetSystemMetrics(0)
            height = user32.GetSystemMetrics(1)
            logger.info(f"Screen resolution: {width}x{height}")
            return (width, height)
        except Exception as e:
            logger.error(f"Error getting screen resolution: {e}")
            logger.error(traceback.format_exc())
            return (1920, 1080)  # Default fallback
    
    def move_mouse_to(self, x, y) -> bool:
        """
        Move the mouse to specific coordinates.
        
        Args:
            x: X coordinate
            y: Y coordinate
            
        Returns:
            True if successful, False otherwise
        """
        try:
            from pywinauto.mouse import move
            logger.info(f"Moving mouse to coordinates: ({x}, {y})")
            move(coords=(x, y))
            return True
        except Exception as e:
            logger.error(f"Failed to move mouse: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def click_at_coordinates(self, x, y, button='left') -> bool:
        """
        Click at specific coordinates.
        
        Args:
            x: X coordinate
            y: Y coordinate
            button: Mouse button ('left', 'right', 'middle')
            
        Returns:
            True if successful, False otherwise
        """
        try:
            from pywinauto.mouse import click
            logger.info(f"Clicking {button} button at coordinates: ({x}, {y})")
            click(button=button, coords=(x, y))
            return True
        except Exception as e:
            logger.error(f"Failed to click at coordinates: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def double_click_at_coordinates(self, x, y) -> bool:
        """
        Double-click at specific coordinates.
        
        Args:
            x: X coordinate
            y: Y coordinate
            
        Returns:
            True if successful, False otherwise
        """
        try:
            from pywinauto.mouse import double_click
            logger.info(f"Double-clicking at coordinates: ({x}, {y})")
            double_click(coords=(x, y))
            return True
        except Exception as e:
            logger.error(f"Failed to double-click at coordinates: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def right_click_at_coordinates(self, x, y) -> bool:
        """
        Right-click at specific coordinates.
        
        Args:
            x: X coordinate
            y: Y coordinate
            
        Returns:
            True if successful, False otherwise
        """
        return self.click_at_coordinates(x, y, button='right')
    
    def send_keys(self, keys) -> bool:
        """
        Send keystrokes to the active window.
        
        Args:
            keys: Keystrokes to send
            
        Returns:
            True if successful, False otherwise
        """
        try:
            from pywinauto.keyboard import send_keys
            logger.info(f"Sending keys: {keys}")
            send_keys(keys)
            return True
        except Exception as e:
            logger.error(f"Failed to send keys: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def press_key(self, key) -> bool:
        """
        Press a key.
        
        Args:
            key: Key to press
            
        Returns:
            True if successful, False otherwise
        """
        try:
            from pywinauto.keyboard import send_keys
            logger.info(f"Pressing key: {key}")
            send_keys(f"{{{key} down}}")
            return True
        except Exception as e:
            logger.error(f"Failed to press key: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def release_key(self, key) -> bool:
        """
        Release a key.
        
        Args:
            key: Key to release
            
        Returns:
            True if successful, False otherwise
        """
        try:
            from pywinauto.keyboard import send_keys
            logger.info(f"Releasing key: {key}")
            send_keys(f"{{{key} up}}")
            return True
        except Exception as e:
            logger.error(f"Failed to release key: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def press_and_release_key(self, key) -> bool:
        """
        Press and release a key.
        
        Args:
            key: Key to press and release
            
        Returns:
            True if successful, False otherwise
        """
        try:
            from pywinauto.keyboard import send_keys
            logger.info(f"Pressing and releasing key: {key}")
            send_keys(f"{{{key}}}")
            return True
        except Exception as e:
            logger.error(f"Failed to press and release key: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def press_key_combination(self, *keys) -> bool:
        """
        Press a combination of keys.
        
        Args:
            *keys: Keys to press
            
        Returns:
            True if successful, False otherwise
        """
        try:
            from pywinauto.keyboard import send_keys
            key_string = '+'.join(keys)
            logger.info(f"Pressing key combination: {key_string}")
            send_keys(f"{{{key_string}}}")
            return True
        except Exception as e:
            logger.error(f"Failed to press key combination: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def wait_for_input_idle(self, timeout=None) -> bool:
        """
        Wait for the application to be idle (not processing input).
        
        Args:
            timeout: Timeout in seconds
            
        Returns:
            True if application is idle, False otherwise
        """
        timeout = timeout or self.default_timeout
        
        try:
            if self.app:
                logger.info(f"Waiting for application to be idle (timeout: {timeout}s)")
                self.app.wait_cpu_usage_lower(threshold=5, timeout=timeout)
                return True
            else:
                logger.warning("No app connection available")
                return False
        except Exception as e:
            logger.error(f"Error waiting for input idle: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def wait_for_process_memory_stable(self, timeout=None, interval=1.0, tolerance=1024*1024) -> bool:
        """
        Wait for the application's memory usage to stabilize.
        
        Args:
            timeout: Timeout in seconds
            interval: Check interval in seconds
            tolerance: Memory difference tolerance in bytes
            
        Returns:
            True if memory usage stabilized, False otherwise
        """
        timeout = timeout or self.default_timeout
        
        try:
            import psutil
            
            if not self.app_pid:
                logger.warning("No application PID available")
                return False
            
            logger.info(f"Waiting for process memory to stabilize (timeout: {timeout}s)")
            
            # Get process
            process = psutil.Process(self.app_pid)
            
            start_time = time.time()
            last_memory = process.memory_info().rss
            
            while time.time() - start_time < timeout:
                time.sleep(interval)
                current_memory = process.memory_info().rss
                
                # Check if memory usage has stabilized
                if abs(current_memory - last_memory) < tolerance:
                    logger.info(f"Process memory has stabilized at {current_memory / (1024*1024):.2f} MB")
                    return True
                
                last_memory = current_memory
            
            logger.warning(f"Process memory did not stabilize after {timeout} seconds")
            return False
            
        except ImportError:
            logger.warning("psutil not available, cannot check memory usage")
            return True
        except Exception as e:
            logger.error(f"Error waiting for process memory to stabilize: {e}")
            logger.error(traceback.format_exc())
            return False
    def get_process_memory_usage(self) -> Optional[int]:
        """
        Get the memory usage of the application process.
        
        Returns:
            Memory usage in bytes or None if failed
        """
        try:
            import psutil
            
            if not self.app_pid:
                logger.warning("No application PID available")
                return None
            
            # Get process
            process = psutil.Process(self.app_pid)
            memory = process.memory_info().rss
            
            logger.info(f"Process memory usage: {memory / (1024*1024):.2f} MB")
            return memory
            
        except ImportError:
            logger.warning("psutil not available, cannot check memory usage")
            return None
        except Exception as e:
            logger.error(f"Error getting process memory usage: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_process_cpu_usage(self) -> Optional[float]:
        """
        Get the CPU usage of the application process.
        
        Returns:
            CPU usage as a percentage or None if failed
        """
        try:
            import psutil
            
            if not self.app_pid:
                logger.warning("No application PID available")
                return None
            
            # Get process
            process = psutil.Process(self.app_pid)
            cpu_usage = process.cpu_percent(interval=0.1)
            
            logger.info(f"Process CPU usage: {cpu_usage:.2f}%")
            return cpu_usage
            
        except ImportError:
            logger.warning("psutil not available, cannot check CPU usage")
            return None
        except Exception as e:
            logger.error(f"Error getting process CPU usage: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def wait_for_cpu_usage_below(self, threshold=5.0, timeout=None) -> bool:
        """
        Wait for the CPU usage to go below a threshold.
        
        Args:
            threshold: CPU usage threshold in percentage
            timeout: Timeout in seconds
            
        Returns:
            True if CPU usage went below threshold, False otherwise
        """
        timeout = timeout or self.default_timeout
        
        try:
            import psutil
            
            if not self.app_pid:
                logger.warning("No application PID available")
                return False
            
            logger.info(f"Waiting for CPU usage to go below {threshold}% (timeout: {timeout}s)")
            
            # Get process
            process = psutil.Process(self.app_pid)
            
            start_time = time.time()
            
            while time.time() - start_time < timeout:
                cpu_usage = process.cpu_percent(interval=0.5)
                
                if cpu_usage < threshold:
                    logger.info(f"CPU usage is now below threshold: {cpu_usage:.2f}%")
                    return True
            
            logger.warning(f"CPU usage did not go below {threshold}% after {timeout} seconds")
            return False
            
        except ImportError:
            logger.warning("psutil not available, cannot check CPU usage")
            return True
        except Exception as e:
            logger.error(f"Error waiting for CPU usage to go below threshold: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_process_responding(self) -> bool:
        """
        Check if the application process is responding.
        
        Returns:
            True if responding, False otherwise
        """
        try:
            import psutil
            
            if not self.app_pid:
                logger.warning("No application PID available")
                return False
            
            # Get process
            process = psutil.Process(self.app_pid)
            
            # Check if process is running
            if process.status() == psutil.STATUS_ZOMBIE:
                logger.warning("Process is in zombie state")
                return False
            
            # Try to get process info as a simple check
            process.cpu_percent()
            
            logger.debug("Process is responding")
            return True
            
        except ImportError:
            logger.warning("psutil not available, using alternative method")
            
            # Alternative method using pywinauto
            if self.app:
                try:
                    return self.app.is_process_running()
                except Exception:
                    return False
            else:
                logger.warning("No app connection available")
                return False
        except Exception as e:
            logger.error(f"Error checking if process is responding: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def kill_process(self) -> bool:
        """
        Kill the application process.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            if self.app:
                logger.info("Killing application process")
                self.app.kill()
                self.app = None
                self.app_pid = None
                return True
            elif self.app_pid:
                import psutil
                logger.info(f"Killing process with PID: {self.app_pid}")
                process = psutil.Process(self.app_pid)
                process.kill()
                self.app_pid = None
                return True
            else:
                logger.warning("No app connection or PID available")
                return False
        except Exception as e:
            logger.error(f"Failed to kill process: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def terminate_process(self) -> bool:
        """
        Terminate the application process (more graceful than kill).
        
        Returns:
            True if successful, False otherwise
        """
        try:
            if self.app:
                logger.info("Terminating application process")
                self.app.kill(soft=True)
                self.app = None
                self.app_pid = None
                return True
            elif self.app_pid:
                import psutil
                logger.info(f"Terminating process with PID: {self.app_pid}")
                process = psutil.Process(self.app_pid)
                process.terminate()
                self.app_pid = None
                return True
            else:
                logger.warning("No app connection or PID available")
                return False
        except Exception as e:
            logger.error(f"Failed to terminate process: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_in_foreground(self, window) -> bool:
        """
        Check if a window is in the foreground.
        
        Args:
            window: Window object
            
        Returns:
            True if in foreground, False otherwise
        """
        try:
            return window.is_active()
        except Exception as e:
            logger.error(f"Error checking if window is in foreground: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def wait_for_window_in_foreground(self, window, timeout=None) -> bool:
        """
        Wait for a window to be in the foreground.
        
        Args:
            window: Window object
            timeout: Timeout in seconds
            
        Returns:
            True if window is in foreground, False otherwise
        """
        timeout = timeout or self.window_timeout
        
        logger.info("Waiting for window to be in foreground")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            if self.is_window_in_foreground(window):
                logger.info("Window is now in foreground")
                return True
            time.sleep(0.5)
        
        logger.warning(f"Window did not come to foreground after {timeout} seconds")
        return False
    
    def wait_for_window_not_busy(self, window, timeout=None) -> bool:
        """
        Wait for a window to not be busy.
        
        Args:
            window: Window object
            timeout: Timeout in seconds
            
        Returns:
            True if window is not busy, False otherwise
        """
        timeout = timeout or self.window_timeout
        
        logger.info("Waiting for window to not be busy")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # Try to interact with the window as a simple check
                window.set_focus()
                logger.info("Window is not busy")
                return True
            except Exception:
                time.sleep(0.5)
        
        logger.warning(f"Window is still busy after {timeout} seconds")
        return False
    
    def capture_window_to_file(self, window, file_path) -> bool:
        """
        Capture a window to an image file.
        
        Args:
            window: Window object
            file_path: Path to save the image
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Capturing window to file: {file_path}")
            
            # Ensure directory exists
            directory = os.path.dirname(file_path)
            if directory and not os.path.exists(directory):
                os.makedirs(directory)
            
            # Capture window
            window.capture_as_image().save(file_path)
            
            logger.info(f"Window captured to file: {file_path}")
            return True
        except Exception as e:
            logger.error(f"Failed to capture window to file: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def get_window_text_elements(self, window) -> List[Any]:
        """
        Get all text elements in a window.
        
        Args:
            window: Window object
            
        Returns:
            List of text elements
        """
        try:
            text_elements = []
            
            # Try different control types that might contain text
            for control_type in ["Text", "Static", "Label", "Edit", "Button"]:
                try:
                    elements = window.children(control_type=control_type)
                    text_elements.extend(elements)
                except Exception:
                    pass
            
            logger.info(f"Found {len(text_elements)} text elements")
            return text_elements
        except Exception as e:
            logger.error(f"Error getting window text elements: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_all_text_from_window(self, window) -> str:
        """
        Get all text from a window.
        
        Args:
            window: Window object
            
        Returns:
            Combined text from all text elements
        """
        try:
            text_elements = self.get_window_text_elements(window)
            
            texts = []
            for element in text_elements:
                try:
                    if hasattr(element, 'window_text'):
                        text = element.window_text()
                        if text:
                            texts.append(text)
                except Exception:
                    pass
            
            combined_text = "\n".join(texts)
            logger.info(f"Got all text from window: {len(combined_text)} characters")
            return combined_text
        except Exception as e:
            logger.error(f"Error getting all text from window: {e}")
            logger.error(traceback.format_exc())
            return ""
    def find_element_by_text(self, window, text, partial_match=False, case_sensitive=True) -> Optional[Any]:
        """
        Find an element by its text.
        
        Args:
            window: Window object to search in
            text: Text to search for
            partial_match: Whether to allow partial matches
            case_sensitive: Whether to use case-sensitive matching
            
        Returns:
            Element if found, None otherwise
        """
        try:
            logger.info(f"Finding element with text: '{text}'")
            
            # Get all descendants
            descendants = window.descendants()
            
            for element in descendants:
                try:
                    if hasattr(element, 'window_text'):
                        element_text = element.window_text()
                        
                        if not case_sensitive:
                            element_text = element_text.lower()
                            search_text = text.lower()
                        else:
                            search_text = text
                        
                        if (partial_match and search_text in element_text) or (not partial_match and search_text == element_text):
                            logger.info(f"Found element with text: '{element_text}'")
                            return element
                except Exception:
                    continue
            
            logger.warning(f"Element with text '{text}' not found")
            return None
        except Exception as e:
            logger.error(f"Error finding element by text: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def find_all_elements_by_text(self, window, text, partial_match=False, case_sensitive=True) -> List[Any]:
        """
        Find all elements by their text.
        
        Args:
            window: Window object to search in
            text: Text to search for
            partial_match: Whether to allow partial matches
            case_sensitive: Whether to use case-sensitive matching
            
        Returns:
            List of elements
        """
        try:
            logger.info(f"Finding all elements with text: '{text}'")
            
            # Get all descendants
            descendants = window.descendants()
            
            matching_elements = []
            for element in descendants:
                try:
                    if hasattr(element, 'window_text'):
                        element_text = element.window_text()
                        
                        if not case_sensitive:
                            element_text = element_text.lower()
                            search_text = text.lower()
                        else:
                            search_text = text
                        
                        if (partial_match and search_text in element_text) or (not partial_match and search_text == element_text):
                            matching_elements.append(element)
                except Exception:
                    continue
            
            logger.info(f"Found {len(matching_elements)} elements with text: '{text}'")
            return matching_elements
        except Exception as e:
            logger.error(f"Error finding elements by text: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def find_element_by_automation_id(self, window, automation_id) -> Optional[Any]:
        """
        Find an element by its automation ID.
        
        Args:
            window: Window object to search in
            automation_id: Automation ID to search for
            
        Returns:
            Element if found, None otherwise
        """
        try:
            logger.info(f"Finding element with automation ID: '{automation_id}'")
            
            # Get all descendants
            descendants = window.descendants()
            
            for element in descendants:
                try:
                    if hasattr(element, 'automation_id'):
                        if callable(element.automation_id):
                            element_id = element.automation_id()
                        else:
                            element_id = element.automation_id
                        
                        if element_id == automation_id:
                            logger.info(f"Found element with automation ID: '{automation_id}'")
                            return element
                except Exception:
                    continue
            
            logger.warning(f"Element with automation ID '{automation_id}' not found")
            return None
        except Exception as e:
            logger.error(f"Error finding element by automation ID: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def find_element_by_control_type(self, window, control_type) -> Optional[Any]:
        """
        Find an element by its control type.
        
        Args:
            window: Window object to search in
            control_type: Control type to search for
            
        Returns:
            Element if found, None otherwise
        """
        try:
            logger.info(f"Finding element with control type: '{control_type}'")
            
            # Get all descendants
            descendants = window.descendants()
            
            for element in descendants:
                try:
                    if hasattr(element, 'control_type'):
                        if callable(element.control_type):
                            element_type = element.control_type()
                        else:
                            element_type = element.control_type
                        
                        if element_type == control_type:
                            logger.info(f"Found element with control type: '{control_type}'")
                            return element
                except Exception:
                    continue
            
            logger.warning(f"Element with control type '{control_type}' not found")
            return None
        except Exception as e:
            logger.error(f"Error finding element by control type: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def find_all_elements_by_control_type(self, window, control_type) -> List[Any]:
        """
        Find all elements by their control type.
        
        Args:
            window: Window object to search in
            control_type: Control type to search for
            
        Returns:
            List of elements
        """
        try:
            logger.info(f"Finding all elements with control type: '{control_type}'")
            
            # Get all descendants
            descendants = window.descendants()
            
            matching_elements = []
            for element in descendants:
                try:
                    if hasattr(element, 'control_type'):
                        if callable(element.control_type):
                            element_type = element.control_type()
                        else:
                            element_type = element.control_type
                        
                        if element_type == control_type:
                            matching_elements.append(element)
                except Exception:
                    continue
            
            logger.info(f"Found {len(matching_elements)} elements with control type: '{control_type}'")
            return matching_elements
        except Exception as e:
            logger.error(f"Error finding elements by control type: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def find_element_by_class_name(self, window, class_name) -> Optional[Any]:
        """
        Find an element by its class name.
        
        Args:
            window: Window object to search in
            class_name: Class name to search for
            
        Returns:
            Element if found, None otherwise
        """
        try:
            logger.info(f"Finding element with class name: '{class_name}'")
            
            # Get all descendants
            descendants = window.descendants()
            
            for element in descendants:
                try:
                    if hasattr(element, 'class_name'):
                        if callable(element.class_name):
                            element_class = element.class_name()
                        else:
                            element_class = element.class_name
                        
                        if element_class == class_name:
                            logger.info(f"Found element with class name: '{class_name}'")
                            return element
                except Exception:
                    continue
            
            logger.warning(f"Element with class name '{class_name}' not found")
            return None
        except Exception as e:
            logger.error(f"Error finding element by class name: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def find_all_elements_by_class_name(self, window, class_name) -> List[Any]:
        """
        Find all elements by their class name.
        
        Args:
            window: Window object to search in
            class_name: Class name to search for
            
        Returns:
            List of elements
        """
        try:
            logger.info(f"Finding all elements with class name: '{class_name}'")
            
            # Get all descendants
            descendants = window.descendants()
            
            matching_elements = []
            for element in descendants:
                try:
                    if hasattr(element, 'class_name'):
                        if callable(element.class_name):
                            element_class = element.class_name()
                        else:
                            element_class = element.class_name
                        
                        if element_class == class_name:
                            matching_elements.append(element)
                except Exception:
                    continue
            
            logger.info(f"Found {len(matching_elements)} elements with class name: '{class_name}'")
            return matching_elements
        except Exception as e:
            logger.error(f"Error finding elements by class name: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def find_child_window(self, parent_window, title=None, class_name=None, control_type=None) -> Optional[Any]:
        """
        Find a child window of a parent window.
        
        Args:
            parent_window: Parent window object
            title: Title of the child window
            class_name: Class name of the child window
            control_type: Control type of the child window
            
        Returns:
            Child window object if found, None otherwise
        """
        try:
            criteria = {}
            if title:
                criteria['title'] = title
            if class_name:
                criteria['class_name'] = class_name
            if control_type:
                criteria['control_type'] = control_type
            
            if not criteria:
                raise ValueError("At least one search criterion must be provided")
            
            logger.info(f"Finding child window with criteria: {criteria}")
            
            child_window = parent_window.child_window(**criteria)
            if child_window.exists():
                logger.info(f"Found child window: '{child_window.window_text() if hasattr(child_window, 'window_text') else 'Unknown'}'")
                return child_window
            else:
                logger.warning("Child window not found")
                return None
        except Exception as e:
            logger.error(f"Error finding child window: {e}")
            logger.error(traceback.format_exc())
            return None
    def find_all_child_windows(self, parent_window, title=None, class_name=None, control_type=None) -> List[Any]:
        """
        Find all child windows of a parent window matching criteria.
        
        Args:
            parent_window: Parent window object
            title: Title of the child windows
            class_name: Class name of the child windows
            control_type: Control type of the child windows
            
        Returns:
            List of child window objects
        """
        try:
            criteria = {}
            if title:
                criteria['title'] = title
            if class_name:
                criteria['class_name'] = class_name
            if control_type:
                criteria['control_type'] = control_type
            
            logger.info(f"Finding all child windows with criteria: {criteria}")
            
            # Get all children
            children = parent_window.children()
            
            # Filter by criteria
            matching_children = []
            for child in children:
                try:
                    match = True
                    if title and hasattr(child, 'window_text') and child.window_text() != title:
                        match = False
                    if class_name and hasattr(child, 'class_name'):
                        if callable(child.class_name):
                            if child.class_name() != class_name:
                                match = False
                        else:
                            if child.class_name != class_name:
                                match = False
                    if control_type and hasattr(child, 'control_type'):
                        if callable(child.control_type):
                            if child.control_type() != control_type:
                                match = False
                        else:
                            if child.control_type != control_type:
                                match = False
                    
                    if match:
                        matching_children.append(child)
                except Exception:
                    continue
            
            logger.info(f"Found {len(matching_children)} matching child windows")
            return matching_children
        except Exception as e:
            logger.error(f"Error finding child windows: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def is_element_clickable(self, element) -> bool:
        """
        Check if an element is clickable (visible and enabled).
        
        Args:
            element: Element to check
            
        Returns:
            True if clickable, False otherwise
        """
        try:
            return element.is_visible() and element.is_enabled()
        except Exception as e:
            logger.debug(f"Error checking if element is clickable: {e}")
            return False
    
    def wait_for_element_clickable(self, element, timeout=None) -> bool:
        """
        Wait for an element to be clickable.
        
        Args:
            element: Element to wait for
            timeout: Timeout in seconds
            
        Returns:
            True if element becomes clickable, False otherwise
        """
        timeout = timeout or self.element_timeout
        
        logger.info("Waiting for element to be clickable")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            if self.is_element_clickable(element):
                logger.info("Element is now clickable")
                return True
            time.sleep(0.5)
        
        logger.warning(f"Element did not become clickable after {timeout} seconds")
        return False
    
    def safe_click_element(self, element, timeout=None) -> bool:
        """
        Safely click an element after ensuring it's clickable.
        
        Args:
            element: Element to click
            timeout: Timeout to wait for element to be clickable
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Wait for element to be clickable
            if not self.wait_for_element_clickable(element, timeout):
                logger.warning("Element is not clickable")
                return False
            
            # Click the element
            logger.info("Clicking element")
            element.click_input()
            return True
        except Exception as e:
            logger.error(f"Failed to safely click element: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def get_element_location(self, element) -> Optional[Tuple[int, int]]:
        """
        Get the screen location of an element.
        
        Args:
            element: Element to get location for
            
        Returns:
            Tuple of (x, y) coordinates or None if failed
        """
        try:
            rect = element.rectangle()
            x = rect.left
            y = rect.top
            logger.debug(f"Element location: ({x}, {y})")
            return (x, y)
        except Exception as e:
            logger.error(f"Error getting element location: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_element_size(self, element) -> Optional[Tuple[int, int]]:
        """
        Get the size of an element.
        
        Args:
            element: Element to get size for
            
        Returns:
            Tuple of (width, height) or None if failed
        """
        try:
            rect = element.rectangle()
            width = rect.width()
            height = rect.height()
            logger.debug(f"Element size: {width}x{height}")
            return (width, height)
        except Exception as e:
            logger.error(f"Error getting element size: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def is_element_checked(self, element) -> Optional[bool]:
        """
        Check if a checkbox or radio button is checked.
        
        Args:
            element: Element to check
            
        Returns:
            True if checked, False if unchecked, None if not applicable or failed
        """
        try:
            if hasattr(element, 'get_toggle_state'):
                state = element.get_toggle_state()
                logger.debug(f"Element toggle state: {state}")
                return state == 1
            elif hasattr(element, 'is_checked'):
                checked = element.is_checked()
                logger.debug(f"Element checked state: {checked}")
                return checked
            else:
                logger.warning("Element doesn't support checked state")
                return None
        except Exception as e:
            logger.error(f"Error checking if element is checked: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def set_checkbox_state(self, checkbox, checked=True) -> bool:
        """
        Set the state of a checkbox.
        
        Args:
            checkbox: Checkbox element
            checked: Whether to check or uncheck
            
        Returns:
            True if successful, False otherwise
        """
        try:
            current_state = self.is_element_checked(checkbox)
            
            # If current state is None, we couldn't determine it
            if current_state is None:
                logger.warning("Could not determine checkbox state")
                return False
            
            # If current state matches desired state, do nothing
            if current_state == checked:
                logger.info(f"Checkbox is already {'checked' if checked else 'unchecked'}")
                return True
            
            # Click to change state
            logger.info(f"Setting checkbox to {'checked' if checked else 'unchecked'}")
            checkbox.click_input()
            
            # Verify the state changed
            new_state = self.is_element_checked(checkbox)
            if new_state == checked:
                logger.info("Checkbox state changed successfully")
                return True
            else:
                logger.warning("Failed to change checkbox state")
                return False
        except Exception as e:
            logger.error(f"Error setting checkbox state: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def select_radio_button(self, radio_button) -> bool:
        """
        Select a radio button.
        
        Args:
            radio_button: Radio button element
            
        Returns:
            True if successful, False otherwise
        """
        try:
            current_state = self.is_element_checked(radio_button)
            
            # If current state is None, we couldn't determine it
            if current_state is None:
                logger.warning("Could not determine radio button state")
                return False
            
            # If already selected, do nothing
            if current_state:
                logger.info("Radio button is already selected")
                return True
            
            # Click to select
            logger.info("Selecting radio button")
            radio_button.click_input()
            
            # Verify the state changed
            new_state = self.is_element_checked(radio_button)
            if new_state:
                logger.info("Radio button selected successfully")
                return True
            else:
                logger.warning("Failed to select radio button")
                return False
        except Exception as e:
            logger.error(f"Error selecting radio button: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_combobox_items(self, combobox) -> List[str]:
        """
        Get all items in a combobox.
        
        Args:
            combobox: Combobox element
            
        Returns:
            List of item texts
        """
        try:
            logger.info("Getting combobox items")
            
            # First click to expand the combobox
            combobox.click_input()
            time.sleep(0.5)
            
            items = []
            
            # Try to find the listbox that appears
            try:
                # Method 1: Look for a listbox as a child of the combobox
                listbox = combobox.child_window(control_type="ListBox")
                if listbox.exists():
                    list_items = listbox.children()
                    for item in list_items:
                        if hasattr(item, 'window_text'):
                            items.append(item.window_text())
            except Exception:
                # Method 2: Look for any listbox that might be the dropdown
                try:
                    listboxes = self.desktop.windows(control_type="ListBox")
                    if listboxes:
                        list_items = listboxes[0].children()
                        for item in list_items:
                            if hasattr(item, 'window_text'):
                                items.append(item.window_text())
                except Exception:
                    pass
            
            # Click again to collapse the combobox
            combobox.click_input()
            
            logger.info(f"Found {len(items)} combobox items")
            return items
        except Exception as e:
            logger.error(f"Error getting combobox items: {e}")
            logger.error(traceback.format_exc())
            
            # Try to collapse the combobox if it's still open
            try:
                combobox.click_input()
            except Exception:
                pass
                
            return []
    
    def get_selected_combobox_item(self, combobox) -> Optional[str]:
        """
        Get the currently selected item in a combobox.
        
        Args:
            combobox: Combobox element
            
        Returns:
            Selected item text or None if failed
        """
        try:
            if hasattr(combobox, 'window_text'):
                text = combobox.window_text()
                logger.info(f"Selected combobox item: {text}")
                return text
            else:
                logger.warning("Combobox doesn't have window_text method")
                return None
        except Exception as e:
            logger.error(f"Error getting selected combobox item: {e}")
            logger.error(traceback.format_exc())
            return None
    def get_table_data(self, table) -> List[List[str]]:
        """
        Get data from a table.
        
        Args:
            table: Table element
            
        Returns:
            List of rows, where each row is a list of cell values
        """
        try:
            logger.info("Getting table data")
            
            # Get all rows
            rows = table.children(control_type="Custom")
            
            table_data = []
            for row in rows:
                row_data = []
                
                # Get cells in the row
                cells = row.children()
                
                for cell in cells:
                    if hasattr(cell, 'window_text'):
                        row_data.append(cell.window_text())
                    else:
                        row_data.append("")
                
                table_data.append(row_data)
            
            logger.info(f"Got data for {len(table_data)} rows")
            return table_data
        except Exception as e:
            logger.error(f"Error getting table data: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_table_headers(self, table) -> List[str]:
        """
        Get headers from a table.
        
        Args:
            table: Table element
            
        Returns:
            List of header texts
        """
        try:
            logger.info("Getting table headers")
            
            # Try to find the header row
            header_row = table.child_window(control_type="Header")
            
            if not header_row.exists():
                logger.warning("Table header row not found")
                return []
            
            # Get header items
            header_items = header_row.children()
            
            headers = []
            for item in header_items:
                if hasattr(item, 'window_text'):
                    headers.append(item.window_text())
                else:
                    headers.append("")
            
            logger.info(f"Got {len(headers)} table headers")
            return headers
        except Exception as e:
            logger.error(f"Error getting table headers: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_table_row_count(self, table) -> int:
        """
        Get the number of rows in a table.
        
        Args:
            table: Table element
            
        Returns:
            Number of rows
        """
        try:
            # Get all rows
            rows = table.children(control_type="Custom")
            count = len(rows)
            
            logger.info(f"Table has {count} rows")
            return count
        except Exception as e:
            logger.error(f"Error getting table row count: {e}")
            logger.error(traceback.format_exc())
            return 0
    
    def get_table_column_count(self, table) -> int:
        """
        Get the number of columns in a table.
        
        Args:
            table: Table element
            
        Returns:
            Number of columns
        """
        try:
            # Try to get the header row first
            try:
                header_row = table.child_window(control_type="Header")
                if header_row.exists():
                    header_items = header_row.children()
                    count = len(header_items)
                    logger.info(f"Table has {count} columns (from headers)")
                    return count
            except Exception:
                pass
            
            # If no header row, try to get the first data row
            try:
                rows = table.children(control_type="Custom")
                if rows:
                    first_row = rows[0]
                    cells = first_row.children()
                    count = len(cells)
                    logger.info(f"Table has {count} columns (from first row)")
                    return count
            except Exception:
                pass
            
            logger.warning("Could not determine table column count")
            return 0
        except Exception as e:
            logger.error(f"Error getting table column count: {e}")
            logger.error(traceback.format_exc())
            return 0
    
    def get_table_cell(self, table, row_index, column_index) -> Optional[str]:
        """
        Get the text of a specific cell in a table.
        
        Args:
            table: Table element
            row_index: Row index (0-based)
            column_index: Column index (0-based)
            
        Returns:
            Cell text or None if failed
        """
        try:
            logger.info(f"Getting table cell at row {row_index}, column {column_index}")
            
            # Get all rows
            rows = table.children(control_type="Custom")
            
            if row_index < 0 or row_index >= len(rows):
                logger.warning(f"Row index {row_index} out of range")
                return None
            
            row = rows[row_index]
            
            # Get cells in the row
            cells = row.children()
            
            if column_index < 0 or column_index >= len(cells):
                logger.warning(f"Column index {column_index} out of range")
                return None
            
            cell = cells[column_index]
            
            if hasattr(cell, 'window_text'):
                text = cell.window_text()
                logger.info(f"Cell text: {text}")
                return text
            else:
                logger.warning("Cell doesn't have window_text method")
                return None
        except Exception as e:
            logger.error(f"Error getting table cell: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def click_table_cell(self, table, row_index, column_index) -> bool:
        """
        Click a specific cell in a table.
        
        Args:
            table: Table element
            row_index: Row index (0-based)
            column_index: Column index (0-based)
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Clicking table cell at row {row_index}, column {column_index}")
            
            # Get all rows
            rows = table.children(control_type="Custom")
            
            if row_index < 0 or row_index >= len(rows):
                logger.warning(f"Row index {row_index} out of range")
                return False
            
            row = rows[row_index]
            
            # Get cells in the row
            cells = row.children()
            
            if column_index < 0 or column_index >= len(cells):
                logger.warning(f"Column index {column_index} out of range")
                return False
            
            cell = cells[column_index]
            
            # Click the cell
            cell.click_input()
            logger.info("Clicked table cell")
            return True
        except Exception as e:
            logger.error(f"Error clicking table cell: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def find_row_by_text(self, table, text, column_index=None, partial_match=False) -> Optional[int]:
        """
        Find a row in a table by text.
        
        Args:
            table: Table element
            text: Text to search for
            column_index: Column index to search in (if None, search all columns)
            partial_match: Whether to allow partial matches
            
        Returns:
            Row index if found, None otherwise
        """
        try:
            logger.info(f"Finding row with text: '{text}'")
            
            # Get all rows
            rows = table.children(control_type="Custom")
            
            for row_index, row in enumerate(rows):
                # Get cells in the row
                cells = row.children()
                
                if column_index is not None:
                    # Search in specific column
                    if column_index < 0 or column_index >= len(cells):
                        logger.warning(f"Column index {column_index} out of range")
                        continue
                    
                    cell = cells[column_index]
                    
                    if hasattr(cell, 'window_text'):
                        cell_text = cell.window_text()
                        
                        if (partial_match and text in cell_text) or (not partial_match and text == cell_text):
                            logger.info(f"Found row {row_index} with text: '{text}'")
                            return row_index
                else:
                    # Search in all columns
                    for cell in cells:
                        if hasattr(cell, 'window_text'):
                            cell_text = cell.window_text()
                            
                            if (partial_match and text in cell_text) or (not partial_match and text == cell_text):
                                logger.info(f"Found row {row_index} with text: '{text}'")
                                return row_index
            
            logger.warning(f"Row with text '{text}' not found")
            return None
        except Exception as e:
            logger.error(f"Error finding row by text: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def find_all_rows_by_text(self, table, text, column_index=None, partial_match=False) -> List[int]:
        """
        Find all rows in a table by text.
        
        Args:
            table: Table element
            text: Text to search for
            column_index: Column index to search in (if None, search all columns)
            partial_match: Whether to allow partial matches
            
        Returns:
            List of row indices
        """
        try:
            logger.info(f"Finding all rows with text: '{text}'")
            
            # Get all rows
            rows = table.children(control_type="Custom")
            
            matching_rows = []
            for row_index, row in enumerate(rows):
                # Get cells in the row
                cells = row.children()
                
                if column_index is not None:
                    # Search in specific column
                    if column_index < 0 or column_index >= len(cells):
                        logger.warning(f"Column index {column_index} out of range")
                        continue
                    
                    cell = cells[column_index]
                    
                    if hasattr(cell, 'window_text'):
                        cell_text = cell.window_text()
                        
                        if (partial_match and text in cell_text) or (not partial_match and text == cell_text):
                            matching_rows.append(row_index)
                else:
                    # Search in all columns
                    for cell in cells:
                        if hasattr(cell, 'window_text'):
                            cell_text = cell.window_text()
                            
                            if (partial_match and text in cell_text) or (not partial_match and text == cell_text):
                                matching_rows.append(row_index)
                                break
            
            logger.info(f"Found {len(matching_rows)} rows with text: '{text}'")
            return matching_rows
        except Exception as e:
            logger.error(f"Error finding rows by text: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_menu_items(self, menu) -> List[Any]:
        """
        Get all items in a menu.
        
        Args:
            menu: Menu element
            
        Returns:
            List of menu item elements
        """
        try:
            logger.info("Getting menu items")
            
            # Get all menu items
            items = menu.children(control_type="MenuItem")
            
            logger.info(f"Found {len(items)} menu items")
            return items
        except Exception as e:
            logger.error(f"Error getting menu items: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_menu_item_texts(self, menu) -> List[str]:
        """
        Get the texts of all items in a menu.
        
        Args:
            menu: Menu element
            
        Returns:
            List of menu item texts
        """
        try:
            logger.info("Getting menu item texts")
            
            # Get all menu items
            items = self.get_menu_items(menu)
            
            texts = []
            for item in items:
                if hasattr(item, 'window_text'):
                    texts.append(item.window_text())
                else:
                    texts.append("")
            
            logger.info(f"Got {len(texts)} menu item texts")
            return texts
        except Exception as e:
            logger.error(f"Error getting menu item texts: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def click_menu_item(self, menu, item_text) -> bool:
        """
        Click a menu item by text.
        
        Args:
            menu: Menu element
            item_text: Text of the menu item to click
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Clicking menu item: '{item_text}'")
            
            # Get all menu items
            items = self.get_menu_items(menu)
            
            for item in items:
                if hasattr(item, 'window_text') and item.window_text() == item_text:
                    item.click_input()
                    logger.info(f"Clicked menu item: '{item_text}'")
                    return True
            
            logger.warning(f"Menu item '{item_text}' not found")
            return False
        except Exception as e:
            logger.error(f"Error clicking menu item: {e}")
            logger.error(traceback.format_exc())
            return False
    def open_context_menu(self, element) -> Optional[Any]:
        """
        Open the context menu for an element.
        
        Args:
            element: Element to right-click
            
        Returns:
            Context menu element or None if failed
        """
        try:
            logger.info("Opening context menu")
            
            # Right-click the element
            element.right_click_input()
            time.sleep(0.5)
            
            # Try to find the context menu
            try:
                # Method 1: Look for a menu as a direct child of the desktop
                context_menu = self.desktop.window(control_type="Menu")
                if context_menu.exists():
                    logger.info("Found context menu (method 1)")
                    return context_menu
            except Exception:
                pass
            
            try:
                # Method 2: Look for any popup menu
                context_menu = self.desktop.window(class_name="#32768")
                if context_menu.exists():
                    logger.info("Found context menu (method 2)")
                    return context_menu
            except Exception:
                pass
            
            logger.warning("Context menu not found")
            return None
        except Exception as e:
            logger.error(f"Error opening context menu: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def click_context_menu_item(self, element, item_text) -> bool:
        """
        Click an item in the context menu of an element.
        
        Args:
            element: Element to right-click
            item_text: Text of the menu item to click
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Clicking context menu item: '{item_text}'")
            
            # Open the context menu
            context_menu = self.open_context_menu(element)
            if not context_menu:
                logger.warning("Could not open context menu")
                return False
            
            # Find and click the menu item
            try:
                menu_item = context_menu.child_window(title=item_text)
                if menu_item.exists():
                    menu_item.click_input()
                    logger.info(f"Clicked context menu item: '{item_text}'")
                    return True
                else:
                    logger.warning(f"Context menu item '{item_text}' not found")
                    return False
            except Exception as e:
                logger.error(f"Error clicking context menu item: {e}")
                logger.error(traceback.format_exc())
                return False
        except Exception as e:
            logger.error(f"Error in click_context_menu_item: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_tree_nodes(self, tree) -> List[Any]:
        """
        Get all top-level nodes in a tree.
        
        Args:
            tree: Tree element
            
        Returns:
            List of tree node elements
        """
        try:
            logger.info("Getting tree nodes")
            
            # Get all top-level tree items
            nodes = tree.children(control_type="TreeItem")
            
            logger.info(f"Found {len(nodes)} tree nodes")
            return nodes
        except Exception as e:
            logger.error(f"Error getting tree nodes: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_tree_node_texts(self, tree) -> List[str]:
        """
        Get the texts of all top-level nodes in a tree.
        
        Args:
            tree: Tree element
            
        Returns:
            List of tree node texts
        """
        try:
            logger.info("Getting tree node texts")
            
            # Get all top-level tree items
            nodes = self.get_tree_nodes(tree)
            
            texts = []
            for node in nodes:
                if hasattr(node, 'window_text'):
                    texts.append(node.window_text())
                else:
                    texts.append("")
            
            logger.info(f"Got {len(texts)} tree node texts")
            return texts
        except Exception as e:
            logger.error(f"Error getting tree node texts: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def expand_tree_node(self, node) -> bool:
        """
        Expand a tree node.
        
        Args:
            node: Tree node element
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Check if node is already expanded
            if hasattr(node, 'is_expanded') and node.is_expanded():
                logger.info("Tree node is already expanded")
                return True
            
            logger.info("Expanding tree node")
            
            # Try to find the expand button
            try:
                expand_button = node.child_window(control_type="Button")
                if expand_button.exists():
                    expand_button.click_input()
                    logger.info("Clicked expand button")
                    return True
            except Exception:
                pass
            
            # If no expand button, try double-clicking the node
            try:
                node.double_click_input()
                logger.info("Double-clicked node to expand")
                return True
            except Exception:
                pass
            
            logger.warning("Could not expand tree node")
            return False
        except Exception as e:
            logger.error(f"Error expanding tree node: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def collapse_tree_node(self, node) -> bool:
        """
        Collapse a tree node.
        
        Args:
            node: Tree node element
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Check if node is already collapsed
            if hasattr(node, 'is_expanded') and not node.is_expanded():
                logger.info("Tree node is already collapsed")
                return True
            
            logger.info("Collapsing tree node")
            
            # Try to find the collapse button
            try:
                collapse_button = node.child_window(control_type="Button")
                if collapse_button.exists():
                    collapse_button.click_input()
                    logger.info("Clicked collapse button")
                    return True
            except Exception:
                pass
            
            # If no collapse button, try double-clicking the node
            try:
                node.double_click_input()
                logger.info("Double-clicked node to collapse")
                return True
            except Exception:
                pass
            
            logger.warning("Could not collapse tree node")
            return False
        except Exception as e:
            logger.error(f"Error collapsing tree node: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_child_tree_nodes(self, node) -> List[Any]:
        """
        Get all child nodes of a tree node.
        
        Args:
            node: Tree node element
            
        Returns:
            List of child tree node elements
        """
        try:
            logger.info("Getting child tree nodes")
            
            # Ensure the node is expanded
            self.expand_tree_node(node)
            
            # Get all child tree items
            children = node.children(control_type="TreeItem")
            
            logger.info(f"Found {len(children)} child tree nodes")
            return children
        except Exception as e:
            logger.error(f"Error getting child tree nodes: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def find_tree_node_by_text(self, tree, text, partial_match=False) -> Optional[Any]:
        """
        Find a tree node by text.
        
        Args:
            tree: Tree element
            text: Text to search for
            partial_match: Whether to allow partial matches
            
        Returns:
            Tree node element if found, None otherwise
        """
        try:
            logger.info(f"Finding tree node with text: '{text}'")
            
            # Get all tree items
            nodes = tree.descendants(control_type="TreeItem")
            
            for node in nodes:
                if hasattr(node, 'window_text'):
                    node_text = node.window_text()
                    
                    if (partial_match and text in node_text) or (not partial_match and text == node_text):
                        logger.info(f"Found tree node with text: '{node_text}'")
                        return node
            
            logger.warning(f"Tree node with text '{text}' not found")
            return None
        except Exception as e:
            logger.error(f"Error finding tree node by text: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def select_tree_node(self, tree, text, partial_match=False) -> bool:
        """
        Select a tree node by text.
        
        Args:
            tree: Tree element
            text: Text of the node to select
            partial_match: Whether to allow partial matches
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Selecting tree node with text: '{text}'")
            
            # Find the node
            node = self.find_tree_node_by_text(tree, text, partial_match)
            if not node:
                logger.warning(f"Tree node with text '{text}' not found")
                return False
            
            # Click the node to select it
            node.click_input()
            logger.info(f"Selected tree node with text: '{text}'")
            return True
        except Exception as e:
            logger.error(f"Error selecting tree node: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def navigate_to_tree_node(self, tree, path) -> Optional[Any]:
        """
        Navigate to a tree node by path.
        
        Args:
            tree: Tree element
            path: List of node texts to navigate through
            
        Returns:
            Final tree node element if found, None otherwise
        """
        try:
            logger.info(f"Navigating to tree node with path: {path}")
            
            current_element = tree
            
            for i, node_text in enumerate(path):
                # Find the node at the current level
                node = self.find_tree_node_by_text(current_element, node_text)
                if not node:
                    logger.warning(f"Tree node '{node_text}' not found at level {i}")
                    return None
                
                # If this is the last node in the path, return it
                if i == len(path) - 1:
                    logger.info(f"Found final tree node: '{node_text}'")
                    return node
                
                # Otherwise, expand this node and continue to the next level
                self.expand_tree_node(node)
                current_element = node
            
            logger.warning("Path is empty")
            return None
        except Exception as e:
            logger.error(f"Error navigating to tree node: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def select_tab(self, tab_control, tab_text) -> bool:
        """
        Select a tab by text.
        
        Args:
            tab_control: Tab control element
            tab_text: Text of the tab to select
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Selecting tab with text: '{tab_text}'")
            
            # Find the tab item
            tab_item = tab_control.child_window(title=tab_text, control_type="TabItem")
            
            if not tab_item.exists():
                logger.warning(f"Tab with text '{tab_text}' not found")
                return False
            
            # Click the tab to select it
            tab_item.click_input()
            logger.info(f"Selected tab with text: '{tab_text}'")
            return True
        except Exception as e:
            logger.error(f"Error selecting tab: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_tab_items(self, tab_control) -> List[Any]:
        """
        Get all tab items in a tab control.
        
        Args:
            tab_control: Tab control element
            
        Returns:
            List of tab item elements
        """
        try:
            logger.info("Getting tab items")
            
            # Get all tab items
            items = tab_control.children(control_type="TabItem")
            
            logger.info(f"Found {len(items)} tab items")
            return items
        except Exception as e:
            logger.error(f"Error getting tab items: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_tab_item_texts(self, tab_control) -> List[str]:
        """
        Get the texts of all tab items in a tab control.
        
        Args:
            tab_control: Tab control element
            
        Returns:
            List of tab item texts
        """
        try:
            logger.info("Getting tab item texts")
            
            # Get all tab items
            items = self.get_tab_items(tab_control)
            
            texts = []
            for item in items:
                if hasattr(item, 'window_text'):
                    texts.append(item.window_text())
                else:
                    texts.append("")
            
            logger.info(f"Got {len(texts)} tab item texts")
            return texts
        except Exception as e:
            logger.error(f"Error getting tab item texts: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_selected_tab(self, tab_control) -> Optional[Any]:
        """
        Get the currently selected tab in a tab control.
        
        Args:
            tab_control: Tab control element
            
        Returns:
            Selected tab item element or None if not found
        """
        try:
            logger.info("Getting selected tab")
            
            # Get all tab items
            items = self.get_tab_items(tab_control)
            
            for item in items:
                if hasattr(item, 'is_selected') and item.is_selected():
                    logger.info(f"Found selected tab: '{item.window_text() if hasattr(item, 'window_text') else 'Unknown'}'")
                    return item
            
            logger.warning("No selected tab found")
            return None
        except Exception as e:
            logger.error(f"Error getting selected tab: {e}")
            logger.error(traceback.format_exc())
            return None
    def get_selected_tab_text(self, tab_control) -> Optional[str]:
        """
        Get the text of the currently selected tab in a tab control.
        
        Args:
            tab_control: Tab control element
            
        Returns:
            Text of the selected tab or None if not found
        """
        try:
            selected_tab = self.get_selected_tab(tab_control)
            if selected_tab and hasattr(selected_tab, 'window_text'):
                text = selected_tab.window_text()
                logger.info(f"Selected tab text: {text}")
                return text
            else:
                logger.warning("Could not get selected tab text")
                return None
        except Exception as e:
            logger.error(f"Error getting selected tab text: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_slider_value(self, slider) -> Optional[int]:
        """
        Get the value of a slider.
        
        Args:
            slider: Slider element
            
        Returns:
            Slider value or None if failed
        """
        try:
            if hasattr(slider, 'get_value'):
                value = slider.get_value()
                logger.info(f"Slider value: {value}")
                return value
            else:
                logger.warning("Slider doesn't have get_value method")
                return None
        except Exception as e:
            logger.error(f"Error getting slider value: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def set_slider_value(self, slider, value) -> bool:
        """
        Set the value of a slider.
        
        Args:
            slider: Slider element
            value: Value to set
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Setting slider value to {value}")
            
            if hasattr(slider, 'set_value'):
                slider.set_value(value)
                logger.info(f"Set slider value to {value}")
                return True
            else:
                logger.warning("Slider doesn't have set_value method")
                return False
        except Exception as e:
            logger.error(f"Error setting slider value: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_progress_bar_value(self, progress_bar) -> Optional[int]:
        """
        Get the value of a progress bar.
        
        Args:
            progress_bar: Progress bar element
            
        Returns:
            Progress bar value or None if failed
        """
        try:
            if hasattr(progress_bar, 'get_value'):
                value = progress_bar.get_value()
                logger.info(f"Progress bar value: {value}")
                return value
            else:
                logger.warning("Progress bar doesn't have get_value method")
                return None
        except Exception as e:
            logger.error(f"Error getting progress bar value: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def wait_for_progress_bar_complete(self, progress_bar, timeout=None) -> bool:
        """
        Wait for a progress bar to complete.
        
        Args:
            progress_bar: Progress bar element
            timeout: Timeout in seconds
            
        Returns:
            True if progress bar completes, False otherwise
        """
        timeout = timeout or self.default_timeout
        
        logger.info("Waiting for progress bar to complete")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            value = self.get_progress_bar_value(progress_bar)
            if value is None:
                logger.warning("Could not get progress bar value")
                return False
            
            if value >= 100:
                logger.info("Progress bar completed")
                return True
            
            time.sleep(0.5)
        
        logger.warning(f"Progress bar did not complete after {timeout} seconds")
        return False
    
    def get_spinner_value(self, spinner) -> Optional[int]:
        """
        Get the value of a spinner.
        
        Args:
            spinner: Spinner element
            
        Returns:
            Spinner value or None if failed
        """
        try:
            if hasattr(spinner, 'get_value'):
                value = spinner.get_value()
                logger.info(f"Spinner value: {value}")
                return value
            elif hasattr(spinner, 'window_text'):
                text = spinner.window_text()
                try:
                    value = int(text)
                    logger.info(f"Spinner value: {value}")
                    return value
                except ValueError:
                    logger.warning(f"Could not convert spinner text '{text}' to integer")
                    return None
            else:
                logger.warning("Spinner doesn't have get_value or window_text method")
                return None
        except Exception as e:
            logger.error(f"Error getting spinner value: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def set_spinner_value(self, spinner, value) -> bool:
        """
        Set the value of a spinner.
        
        Args:
            spinner: Spinner element
            value: Value to set
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Setting spinner value to {value}")
            
            if hasattr(spinner, 'set_value'):
                spinner.set_value(value)
                logger.info(f"Set spinner value to {value}")
                return True
            elif hasattr(spinner, 'set_text'):
                spinner.set_text(str(value))
                logger.info(f"Set spinner text to {value}")
                return True
            else:
                logger.warning("Spinner doesn't have set_value or set_text method")
                return False
        except Exception as e:
            logger.error(f"Error setting spinner value: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def increment_spinner(self, spinner) -> bool:
        """
        Increment a spinner.
        
        Args:
            spinner: Spinner element
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info("Incrementing spinner")
            
            # Try to find the up button
            try:
                up_button = spinner.child_window(control_type="Button")
                if up_button.exists():
                    up_button.click_input()
                    logger.info("Clicked up button")
                    return True
            except Exception:
                pass
            
            # If no up button, try to increment the value directly
            current_value = self.get_spinner_value(spinner)
            if current_value is not None:
                return self.set_spinner_value(spinner, current_value + 1)
            
            logger.warning("Could not increment spinner")
            return False
        except Exception as e:
            logger.error(f"Error incrementing spinner: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def decrement_spinner(self, spinner) -> bool:
        """
        Decrement a spinner.
        
        Args:
            spinner: Spinner element
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info("Decrementing spinner")
            
            # Try to find the down button
            try:
                buttons = spinner.children(control_type="Button")
                if len(buttons) > 1:
                    down_button = buttons[1]
                    down_button.click_input()
                    logger.info("Clicked down button")
                    return True
            except Exception:
                pass
            
            # If no down button, try to decrement the value directly
            current_value = self.get_spinner_value(spinner)
            if current_value is not None:
                return self.set_spinner_value(spinner, current_value - 1)
            
            logger.warning("Could not decrement spinner")
            return False
        except Exception as e:
            logger.error(f"Error decrementing spinner: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_date_picker_value(self, date_picker) -> Optional[str]:
        """
        Get the value of a date picker.
        
        Args:
            date_picker: Date picker element
            
        Returns:
            Date string or None if failed
        """
        try:
            if hasattr(date_picker, 'window_text'):
                text = date_picker.window_text()
                logger.info(f"Date picker value: {text}")
                return text
            else:
                logger.warning("Date picker doesn't have window_text method")
                return None
        except Exception as e:
            logger.error(f"Error getting date picker value: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def set_date_picker_value(self, date_picker, date_str) -> bool:
        """
        Set the value of a date picker.
        
        Args:
            date_picker: Date picker element
            date_str: Date string in the format expected by the date picker
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Setting date picker value to {date_str}")
            
            # Click to open the date picker
            date_picker.click_input()
            time.sleep(0.5)
            
            # Try to find an edit field to enter the date
            try:
                edit_field = date_picker.child_window(control_type="Edit")
                if edit_field.exists():
                    edit_field.set_text(date_str)
                    
                    # Press Enter to confirm
                    from pywinauto.keyboard import send_keys
                    send_keys("{ENTER}")
                    
                    logger.info(f"Set date picker value to {date_str}")
                    return True
            except Exception:
                pass
            
            logger.warning("Could not set date picker value")
            return False
        except Exception as e:
            logger.error(f"Error setting date picker value: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_list_items(self, list_control) -> List[Any]:
        """
        Get all items in a list control.
        
        Args:
            list_control: List control element
            
        Returns:
            List of list item elements
        """
        try:
            logger.info("Getting list items")
            
            # Get all list items
            items = list_control.children(control_type="ListItem")
            
            logger.info(f"Found {len(items)} list items")
            return items
        except Exception as e:
            logger.error(f"Error getting list items: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_list_item_texts(self, list_control) -> List[str]:
        """
        Get the texts of all items in a list control.
        
        Args:
            list_control: List control element
            
        Returns:
            List of list item texts
        """
        try:
            logger.info("Getting list item texts")
            
            # Get all list items
            items = self.get_list_items(list_control)
            
            texts = []
            for item in items:
                if hasattr(item, 'window_text'):
                    texts.append(item.window_text())
                else:
                    texts.append("")
            
            logger.info(f"Got {len(texts)} list item texts")
            return texts
        except Exception as e:
            logger.error(f"Error getting list item texts: {e}")
            logger.error(traceback.format_exc())
            return []
    def select_list_item(self, list_control, item_text) -> bool:
        """
        Select an item in a list control by text.
        
        Args:
            list_control: List control element
            item_text: Text of the item to select
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Selecting list item with text: '{item_text}'")
            
            # Find the list item
            list_item = list_control.child_window(title=item_text, control_type="ListItem")
            
            if not list_item.exists():
                logger.warning(f"List item with text '{item_text}' not found")
                return False
            
            # Click the list item to select it
            list_item.click_input()
            logger.info(f"Selected list item with text: '{item_text}'")
            return True
        except Exception as e:
            logger.error(f"Error selecting list item: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_selected_list_item(self, list_control) -> Optional[Any]:
        """
        Get the currently selected item in a list control.
        
        Args:
            list_control: List control element
            
        Returns:
            Selected list item element or None if not found
        """
        try:
            logger.info("Getting selected list item")
            
            # Get all list items
            items = self.get_list_items(list_control)
            
            for item in items:
                if hasattr(item, 'is_selected') and item.is_selected():
                    logger.info(f"Found selected list item: '{item.window_text() if hasattr(item, 'window_text') else 'Unknown'}'")
                    return item
            
            logger.warning("No selected list item found")
            return None
        except Exception as e:
            logger.error(f"Error getting selected list item: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_selected_list_item_text(self, list_control) -> Optional[str]:
        """
        Get the text of the currently selected item in a list control.
        
        Args:
            list_control: List control element
            
        Returns:
            Text of the selected list item or None if not found
        """
        try:
            selected_item = self.get_selected_list_item(list_control)
            if selected_item and hasattr(selected_item, 'window_text'):
                text = selected_item.window_text()
                logger.info(f"Selected list item text: {text}")
                return text
            else:
                logger.warning("Could not get selected list item text")
                return None
        except Exception as e:
            logger.error(f"Error getting selected list item text: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def select_multiple_list_items(self, list_control, item_texts) -> bool:
        """
        Select multiple items in a list control by text.
        
        Args:
            list_control: List control element
            item_texts: List of texts of the items to select
            
        Returns:
            True if all items were selected, False otherwise
        """
        try:
            logger.info(f"Selecting multiple list items: {item_texts}")
            
            # Clear current selection first
            list_control.click_input()
            
            # Hold Ctrl key for multiple selection
            from pywinauto.keyboard import send_keys
            send_keys("{VK_CONTROL down}")
            
            success = True
            for item_text in item_texts:
                # Find the list item
                list_item = list_control.child_window(title=item_text, control_type="ListItem")
                
                if not list_item.exists():
                    logger.warning(f"List item with text '{item_text}' not found")
                    success = False
                    continue
                
                # Click the list item to select it
                list_item.click_input()
                logger.info(f"Selected list item with text: '{item_text}'")
            
            # Release Ctrl key
            send_keys("{VK_CONTROL up}")
            
            return success
        except Exception as e:
            logger.error(f"Error selecting multiple list items: {e}")
            logger.error(traceback.format_exc())
            
            # Make sure to release Ctrl key
            try:
                from pywinauto.keyboard import send_keys
                send_keys("{VK_CONTROL up}")
            except Exception:
                pass
                
            return False
    
    def get_selected_list_items(self, list_control) -> List[Any]:
        """
        Get all selected items in a list control.
        
        Args:
            list_control: List control element
            
        Returns:
            List of selected list item elements
        """
        try:
            logger.info("Getting selected list items")
            
            # Get all list items
            items = self.get_list_items(list_control)
            
            selected_items = []
            for item in items:
                if hasattr(item, 'is_selected') and item.is_selected():
                    selected_items.append(item)
            
            logger.info(f"Found {len(selected_items)} selected list items")
            return selected_items
        except Exception as e:
            logger.error(f"Error getting selected list items: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_selected_list_item_texts(self, list_control) -> List[str]:
        """
        Get the texts of all selected items in a list control.
        
        Args:
            list_control: List control element
            
        Returns:
            List of selected list item texts
        """
        try:
            logger.info("Getting selected list item texts")
            
            # Get all selected list items
            items = self.get_selected_list_items(list_control)
            
            texts = []
            for item in items:
                if hasattr(item, 'window_text'):
                    texts.append(item.window_text())
                else:
                    texts.append("")
            
            logger.info(f"Got {len(texts)} selected list item texts")
            return texts
        except Exception as e:
            logger.error(f"Error getting selected list item texts: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def is_dialog_open(self, title=None, class_name="#32770") -> bool:
        """
        Check if a dialog is open.
        
        Args:
            title: Dialog title (optional)
            class_name: Dialog class name (default is "#32770" for standard Windows dialogs)
            
        Returns:
            True if dialog is open, False otherwise
        """
        try:
            criteria = {}
            if title:
                criteria['title'] = title
            if class_name:
                criteria['class_name'] = class_name
                
            logger.info(f"Checking if dialog is open with criteria: {criteria}")
            
            # Use Desktop to find all windows
            desktop = Desktop(backend="uia")
            
            # Try to find the dialog
            try:
                dialog = desktop.window(**criteria)
                if dialog.exists():
                    logger.info(f"Dialog is open: '{dialog.window_text() if hasattr(dialog, 'window_text') else 'Unknown'}'")
                    return True
                else:
                    logger.info("Dialog is not open")
                    return False
            except Exception:
                logger.info("Dialog is not open")
                return False
        except Exception as e:
            logger.error(f"Error checking if dialog is open: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def wait_for_dialog(self, title=None, class_name="#32770", timeout=None) -> Optional[Any]:
        """
        Wait for a dialog to open.
        
        Args:
            title: Dialog title (optional)
            class_name: Dialog class name (default is "#32770" for standard Windows dialogs)
            timeout: Timeout in seconds
            
        Returns:
            Dialog window object if found, None otherwise
        """
        timeout = timeout or self.window_timeout
        
        criteria = {}
        if title:
            criteria['title'] = title
        if class_name:
            criteria['class_name'] = class_name
            
        logger.info(f"Waiting for dialog with criteria: {criteria} (timeout: {timeout}s)")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            # Use Desktop to find all windows
            desktop = Desktop(backend="uia")
            
            # Try to find the dialog
            try:
                dialog = desktop.window(**criteria)
                if dialog.exists():
                    logger.info(f"Dialog found: '{dialog.window_text() if hasattr(dialog, 'window_text') else 'Unknown'}'")
                    return dialog
            except Exception:
                pass
                
            time.sleep(0.5)
        
        logger.warning(f"Dialog not found after {timeout} seconds")
        return None
    
    def close_dialog(self, dialog, button_text="OK") -> bool:
        """
        Close a dialog by clicking a button.
        
        Args:
            dialog: Dialog window object
            button_text: Text of the button to click (default is "OK")
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Closing dialog by clicking '{button_text}' button")
            
            # Find the button
            button = dialog.child_window(title=button_text, control_type="Button")
            
            if not button.exists():
                logger.warning(f"Button with text '{button_text}' not found")
                return False
            
            # Click the button
            button.click_input()
            logger.info(f"Clicked '{button_text}' button")
            
            # Wait for the dialog to close
            start_time = time.time()
            while time.time() - start_time < 5:  # 5 seconds timeout
                try:
                    if not dialog.exists():
                        logger.info("Dialog closed successfully")
                        return True
                except Exception:
                    # If checking fails, assume the dialog is closed
                    logger.info("Dialog closed successfully (exception while checking)")
                    return True
                    
                time.sleep(0.5)
            
            logger.warning("Dialog did not close after clicking button")
            return False
        except Exception as e:
            logger.error(f"Error closing dialog: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def handle_dialog(self, title=None, class_name="#32770", button_text="OK", timeout=None) -> bool:
        """
        Wait for a dialog and close it by clicking a button.
        
        Args:
            title: Dialog title (optional)
            class_name: Dialog class name (default is "#32770" for standard Windows dialogs)
            button_text: Text of the button to click (default is "OK")
            timeout: Timeout in seconds
            
        Returns:
            True if successful, False otherwise
        """
        dialog = self.wait_for_dialog(title, class_name, timeout)
        if dialog:
            return self.close_dialog(dialog, button_text)
        else:
            return False
    def get_dialog_message(self, dialog) -> str:
        """
        Get the message text from a dialog.
        
        Args:
            dialog: Dialog window object
            
        Returns:
            Message text
        """
        try:
            logger.info("Getting dialog message")
            
            # Try to find a static text control
            try:
                static_text = dialog.child_window(control_type="Text")
                if static_text.exists():
                    message = static_text.window_text()
                    logger.info(f"Dialog message: {message}")
                    return message
            except Exception:
                pass
            
            # Try to find a static control
            try:
                static = dialog.child_window(control_type="Static")
                if static.exists():
                    message = static.window_text()
                    logger.info(f"Dialog message: {message}")
                    return message
            except Exception:
                pass
            
            # If no specific text control found, get all text from the dialog
            all_text = self.get_all_text_from_window(dialog)
            if all_text:
                logger.info(f"Dialog message (all text): {all_text}")
                return all_text
            
            # If all else fails, return the dialog title
            if hasattr(dialog, 'window_text'):
                title = dialog.window_text()
                logger.info(f"Dialog message (title): {title}")
                return title
            
            logger.warning("Could not get dialog message")
            return ""
        except Exception as e:
            logger.error(f"Error getting dialog message: {e}")
            logger.error(traceback.format_exc())
            return ""
    
    def handle_dialog_with_message(self, title=None, class_name="#32770", button_text="OK", timeout=None) -> Tuple[bool, str]:
        """
        Wait for a dialog, get its message, and close it by clicking a button.
        
        Args:
            title: Dialog title (optional)
            class_name: Dialog class name (default is "#32770" for standard Windows dialogs)
            button_text: Text of the button to click (default is "OK")
            timeout: Timeout in seconds
            
        Returns:
            Tuple of (success, message)
        """
        dialog = self.wait_for_dialog(title, class_name, timeout)
        if dialog:
            message = self.get_dialog_message(dialog)
            success = self.close_dialog(dialog, button_text)
            return success, message
        else:
            return False, ""
    
    def is_error_dialog_open(self) -> bool:
        """
        Check if an error dialog is open.
        
        Returns:
            True if an error dialog is open, False otherwise
        """
        # Check for common error dialog titles
        for title in ["Error", "Exception", "Critical Error", "Warning"]:
            if self.is_dialog_open(title):
                return True
        
        # Check for dialog with "error" in the title
        try:
            desktop = Desktop(backend="uia")
            for window in desktop.windows():
                try:
                    if hasattr(window, 'window_text') and "error" in window.window_text().lower():
                        return True
                except Exception:
                    pass
        except Exception:
            pass
            
        return False
    
    def handle_error_dialog(self, button_text="OK", timeout=None) -> Tuple[bool, str]:
        """
        Wait for an error dialog, get its message, and close it by clicking a button.
        
        Args:
            button_text: Text of the button to click (default is "OK")
            timeout: Timeout in seconds
            
        Returns:
            Tuple of (success, message)
        """
        # Check for common error dialog titles
        for title in ["Error", "Exception", "Critical Error", "Warning"]:
            dialog = self.wait_for_dialog(title, timeout=1)
            if dialog:
                message = self.get_dialog_message(dialog)
                success = self.close_dialog(dialog, button_text)
                return success, message
        
        # Check for dialog with "error" in the title
        try:
            desktop = Desktop(backend="uia")
            for window in desktop.windows():
                try:
                    if hasattr(window, 'window_text') and "error" in window.window_text().lower():
                        message = self.get_dialog_message(window)
                        success = self.close_dialog(window, button_text)
                        return success, message
                except Exception:
                    pass
        except Exception:
            pass
            
        return False, ""
    
    def wait_for_window_state(self, window, state_check_func, timeout=None) -> bool:
        """
        Wait for a window to reach a specific state.
        
        Args:
            window: Window object
            state_check_func: Function that takes the window as argument and returns True when the desired state is reached
            timeout: Timeout in seconds
            
        Returns:
            True if the window reached the desired state, False otherwise
        """
        timeout = timeout or self.window_timeout
        
        logger.info("Waiting for window to reach desired state")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                if state_check_func(window):
                    logger.info("Window reached desired state")
                    return True
            except Exception as e:
                logger.debug(f"Error checking window state: {e}")
                
            time.sleep(0.5)
        
        logger.warning(f"Window did not reach desired state after {timeout} seconds")
        return False
    
    def wait_for_window_ready(self, window, timeout=None) -> bool:
        """
        Wait for a window to be ready (not busy).
        
        Args:
            window: Window object
            timeout: Timeout in seconds
            
        Returns:
            True if the window is ready, False otherwise
        """
        def is_window_ready(w):
            try:
                # Try to interact with the window as a simple check
                w.set_focus()
                return True
            except Exception:
                return False
                
        return self.wait_for_window_state(window, is_window_ready, timeout)
    
    def wait_for_window_enabled(self, window, timeout=None) -> bool:
        """
        Wait for a window to be enabled.
        
        Args:
            window: Window object
            timeout: Timeout in seconds
            
        Returns:
            True if the window is enabled, False otherwise
        """
        def is_window_enabled(w):
            return w.is_enabled()
                
        return self.wait_for_window_state(window, is_window_enabled, timeout)
    
    def wait_for_window_active(self, window, timeout=None) -> bool:
        """
        Wait for a window to be active (in foreground).
        
        Args:
            window: Window object
            timeout: Timeout in seconds
            
        Returns:
            True if the window is active, False otherwise
        """
        def is_window_active(w):
            return w.is_active()
                
        return self.wait_for_window_state(window, is_window_active, timeout)
    
    def wait_for_window_text(self, window, text, timeout=None) -> bool:
        """
        Wait for a window to have specific text.
        
        Args:
            window: Window object
            text: Text to wait for
            timeout: Timeout in seconds
            
        Returns:
            True if the window has the text, False otherwise
        """
        def has_window_text(w):
            return w.window_text() == text
                
        return self.wait_for_window_state(window, has_window_text, timeout)
    
    def wait_for_window_text_contains(self, window, text, timeout=None) -> bool:
        """
        Wait for a window to have text containing a specific substring.
        
        Args:
            window: Window object
            text: Substring to wait for
            timeout: Timeout in seconds
            
        Returns:
            True if the window text contains the substring, False otherwise
        """
        def window_text_contains(w):
            return text in w.window_text()
                
        return self.wait_for_window_state(window, window_text_contains, timeout)
    
    def wait_for_element_state(self, element, state_check_func, timeout=None) -> bool:
        """
        Wait for an element to reach a specific state.
        
        Args:
            element: Element object
            state_check_func: Function that takes the element as argument and returns True when the desired state is reached
            timeout: Timeout in seconds
            
        Returns:
            True if the element reached the desired state, False otherwise
        """
        timeout = timeout or self.element_timeout
        
        logger.info("Waiting for element to reach desired state")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                if state_check_func(element):
                    logger.info("Element reached desired state")
                    return True
            except Exception as e:
                logger.debug(f"Error checking element state: {e}")
                
            time.sleep(0.5)
        
        logger.warning(f"Element did not reach desired state after {timeout} seconds")
        return False
    
    def wait_for_element_text(self, element, text, timeout=None) -> bool:
        """
        Wait for an element to have specific text.
        
        Args:
            element: Element object
            text: Text to wait for
            timeout: Timeout in seconds
            
        Returns:
            True if the element has the text, False otherwise
        """
        def has_element_text(e):
            return e.window_text() == text
                
        return self.wait_for_element_state(element, has_element_text, timeout)
    
    def wait_for_element_text_contains(self, element, text, timeout=None) -> bool:
        """
        Wait for an element to have text containing a specific substring.
        
        Args:
            element: Element object
            text: Substring to wait for
            timeout: Timeout in seconds
            
        Returns:
            True if the element text contains the substring, False otherwise
        """
        def element_text_contains(e):
            return text in e.window_text()
                
        return self.wait_for_element_state(element, element_text_contains, timeout)
    def wait_for_element_property(self, element, property_name, expected_value, timeout=None) -> bool:
        """
        Wait for an element property to have a specific value.
        
        Args:
            element: Element object
            property_name: Name of the property
            expected_value: Expected value of the property
            timeout: Timeout in seconds
            
        Returns:
            True if the property has the expected value, False otherwise
        """
        def has_property_value(e):
            if hasattr(e, property_name):
                attr = getattr(e, property_name)
                if callable(attr):
                    return attr() == expected_value
                else:
                    return attr == expected_value
            return False
                
        return self.wait_for_element_state(element, has_property_value, timeout)
    
    def wait_for_element_property_contains(self, element, property_name, expected_substring, timeout=None) -> bool:
        """
        Wait for an element property to contain a specific substring.
        
        Args:
            element: Element object
            property_name: Name of the property
            expected_substring: Expected substring in the property value
            timeout: Timeout in seconds
            
        Returns:
            True if the property contains the expected substring, False otherwise
        """
        def property_contains(e):
            if hasattr(e, property_name):
                attr = getattr(e, property_name)
                if callable(attr):
                    value = attr()
                else:
                    value = attr
                return expected_substring in str(value)
            return False
                
        return self.wait_for_element_state(element, property_contains, timeout)
    
    def wait_for_element_exists(self, parent_window, control_type=None, name=None, automation_id=None, class_name=None, timeout=None) -> Optional[Any]:
        """
        Wait for an element to exist.
        
        Args:
            parent_window: Parent window object
            control_type: Control type
            name: Element name
            automation_id: Automation ID
            class_name: Class name
            timeout: Timeout in seconds
            
        Returns:
            Element object if found, None otherwise
        """
        timeout = timeout or self.element_timeout
        
        criteria = {}
        if control_type:
            criteria['control_type'] = control_type
        if name:
            criteria['title'] = name
        if automation_id:
            criteria['automation_id'] = automation_id
        if class_name:
            criteria['class_name'] = class_name
        
        if not criteria:
            raise ValueError("At least one search criterion must be provided")
        
        logger.info(f"Waiting for element to exist with criteria: {criteria}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                element = self.find_element(parent_window, control_type, name, automation_id, class_name, 1)
                if element:
                    logger.info("Element exists")
                    return element
            except Exception as e:
                logger.debug(f"Error checking if element exists: {e}")
                
            time.sleep(0.5)
        
        logger.warning(f"Element does not exist after {timeout} seconds")
        return None
    
    def wait_for_element_not_exists(self, parent_window, control_type=None, name=None, automation_id=None, class_name=None, timeout=None) -> bool:
        """
        Wait for an element to not exist.
        
        Args:
            parent_window: Parent window object
            control_type: Control type
            name: Element name
            automation_id: Automation ID
            class_name: Class name
            timeout: Timeout in seconds
            
        Returns:
            True if the element does not exist, False otherwise
        """
        timeout = timeout or self.element_timeout
        
        criteria = {}
        if control_type:
            criteria['control_type'] = control_type
        if name:
            criteria['title'] = name
        if automation_id:
            criteria['automation_id'] = automation_id
        if class_name:
            criteria['class_name'] = class_name
        
        if not criteria:
            raise ValueError("At least one search criterion must be provided")
        
        logger.info(f"Waiting for element to not exist with criteria: {criteria}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                element = self.find_element(parent_window, control_type, name, automation_id, class_name, 1)
                if not element:
                    logger.info("Element does not exist")
                    return True
            except Exception:
                # If find_element raises an exception, the element probably doesn't exist
                logger.info("Element does not exist (exception while checking)")
                return True
                
            time.sleep(0.5)
        
        logger.warning(f"Element still exists after {timeout} seconds")
        return False
    
    def wait_for_any_element(self, parent_window, criteria_list, timeout=None) -> Optional[Any]:
        """
        Wait for any of several elements to exist.
        
        Args:
            parent_window: Parent window object
            criteria_list: List of criteria dictionaries, each containing control_type, name, automation_id, and/or class_name
            timeout: Timeout in seconds
            
        Returns:
            First element found or None if none found
        """
        timeout = timeout or self.element_timeout
        
        if not criteria_list:
            raise ValueError("Criteria list cannot be empty")
        
        logger.info(f"Waiting for any element from criteria list: {criteria_list}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            for criteria in criteria_list:
                control_type = criteria.get('control_type')
                name = criteria.get('name')
                automation_id = criteria.get('automation_id')
                class_name = criteria.get('class_name')
                
                try:
                    element = self.find_element(parent_window, control_type, name, automation_id, class_name, 1)
                    if element:
                        logger.info(f"Found element with criteria: {criteria}")
                        return element
                except Exception:
                    pass
                    
            time.sleep(0.5)
        
        logger.warning(f"No elements found after {timeout} seconds")
        return None
    
    def wait_for_window_appears(self, title=None, class_name=None, timeout=None) -> Optional[Any]:
        """
        Wait for a window to appear.
        
        Args:
            title: Window title
            class_name: Window class name
            timeout: Timeout in seconds
            
        Returns:
            Window object if found, None otherwise
        """
        timeout = timeout or self.window_timeout
        
        criteria = {}
        if title:
            criteria['title'] = title
        if class_name:
            criteria['class_name'] = class_name
        
        if not criteria:
            raise ValueError("At least one search criterion must be provided")
        
        logger.info(f"Waiting for window to appear with criteria: {criteria}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                window = self.find_window(title=title, class_name=class_name, timeout=1)
                if window:
                    logger.info("Window appeared")
                    return window
            except Exception:
                pass
                
            time.sleep(0.5)
        
        logger.warning(f"Window did not appear after {timeout} seconds")
        return None
    
    def wait_for_window_disappears(self, title=None, class_name=None, timeout=None) -> bool:
        """
        Wait for a window to disappear.
        
        Args:
            title: Window title
            class_name: Window class name
            timeout: Timeout in seconds
            
        Returns:
            True if the window disappeared, False otherwise
        """
        timeout = timeout or self.window_timeout
        
        criteria = {}
        if title:
            criteria['title'] = title
        if class_name:
            criteria['class_name'] = class_name
        
        if not criteria:
            raise ValueError("At least one search criterion must be provided")
        
        logger.info(f"Waiting for window to disappear with criteria: {criteria}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                window = self.find_window(title=title, class_name=class_name, timeout=1)
                if not window:
                    logger.info("Window disappeared")
                    return True
            except Exception:
                # If find_window raises an exception, the window probably doesn't exist
                logger.info("Window disappeared (exception while checking)")
                return True
                
            time.sleep(0.5)
        
        logger.warning(f"Window still exists after {timeout} seconds")
        return False
    
    def wait_for_any_window(self, criteria_list, timeout=None) -> Optional[Any]:
        """
        Wait for any of several windows to appear.
        
        Args:
            criteria_list: List of criteria dictionaries, each containing title and/or class_name
            timeout: Timeout in seconds
            
        Returns:
            First window found or None if none found
        """
        timeout = timeout or self.window_timeout
        
        if not criteria_list:
            raise ValueError("Criteria list cannot be empty")
        
        logger.info(f"Waiting for any window from criteria list: {criteria_list}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            for criteria in criteria_list:
                title = criteria.get('title')
                class_name = criteria.get('class_name')
                
                try:
                    window = self.find_window(title=title, class_name=class_name, timeout=1)
                    if window:
                        logger.info(f"Found window with criteria: {criteria}")
                        return window
                except Exception:
                    pass
                    
            time.sleep(0.5)
        
        logger.warning(f"No windows found after {timeout} seconds")
        return None
    def get_window_children_texts(self, window) -> List[str]:
        """
        Get the texts of all child elements of a window.
        
        Args:
            window: Window object
            
        Returns:
            List of child element texts
        """
        try:
            logger.info("Getting window children texts")
            
            # Get all children
            children = window.children()
            
            texts = []
            for child in children:
                if hasattr(child, 'window_text'):
                    text = child.window_text()
                    if text:
                        texts.append(text)
            
            logger.info(f"Got {len(texts)} window children texts")
            return texts
        except Exception as e:
            logger.error(f"Error getting window children texts: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_window_descendants_texts(self, window) -> List[str]:
        """
        Get the texts of all descendant elements of a window.
        
        Args:
            window: Window object
            
        Returns:
            List of descendant element texts
        """
        try:
            logger.info("Getting window descendants texts")
            
            # Get all descendants
            descendants = window.descendants()
            
            texts = []
            for descendant in descendants:
                if hasattr(descendant, 'window_text'):
                    text = descendant.window_text()
                    if text:
                        texts.append(text)
            
            logger.info(f"Got {len(texts)} window descendants texts")
            return texts
        except Exception as e:
            logger.error(f"Error getting window descendants texts: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def find_element_by_text_in_window(self, window, text, partial_match=False) -> Optional[Any]:
        """
        Find an element by its text in a window.
        
        Args:
            window: Window object
            text: Text to search for
            partial_match: Whether to allow partial matches
            
        Returns:
            Element if found, None otherwise
        """
        try:
            logger.info(f"Finding element with text '{text}' in window")
            
            # Get all descendants
            descendants = window.descendants()
            
            for descendant in descendants:
                if hasattr(descendant, 'window_text'):
                    element_text = descendant.window_text()
                    
                    if (partial_match and text in element_text) or (not partial_match and text == element_text):
                        logger.info(f"Found element with text: '{element_text}'")
                        return descendant
            
            logger.warning(f"Element with text '{text}' not found in window")
            return None
        except Exception as e:
            logger.error(f"Error finding element by text in window: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def find_all_elements_by_text_in_window(self, window, text, partial_match=False) -> List[Any]:
        """
        Find all elements by their text in a window.
        
        Args:
            window: Window object
            text: Text to search for
            partial_match: Whether to allow partial matches
            
        Returns:
            List of elements
        """
        try:
            logger.info(f"Finding all elements with text '{text}' in window")
            
            # Get all descendants
            descendants = window.descendants()
            
            matching_elements = []
            for descendant in descendants:
                if hasattr(descendant, 'window_text'):
                    element_text = descendant.window_text()
                    
                    if (partial_match and text in element_text) or (not partial_match and text == element_text):
                        matching_elements.append(descendant)
            
            logger.info(f"Found {len(matching_elements)} elements with text '{text}' in window")
            return matching_elements
        except Exception as e:
            logger.error(f"Error finding elements by text in window: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def click_element_by_text(self, window, text, partial_match=False) -> bool:
        """
        Click an element by its text.
        
        Args:
            window: Window object
            text: Text of the element to click
            partial_match: Whether to allow partial matches
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Clicking element with text '{text}'")
            
            # Find the element
            element = self.find_element_by_text_in_window(window, text, partial_match)
            if not element:
                logger.warning(f"Element with text '{text}' not found")
                return False
            
            # Click the element
            element.click_input()
            logger.info(f"Clicked element with text '{text}'")
            return True
        except Exception as e:
            logger.error(f"Error clicking element by text: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_window_title(self, window) -> str:
        """
        Get the title of a window.
        
        Args:
            window: Window object
            
        Returns:
            Window title
        """
        try:
            if hasattr(window, 'window_text'):
                title = window.window_text()
                logger.info(f"Window title: {title}")
                return title
            else:
                logger.warning("Window doesn't have window_text method")
                return ""
        except Exception as e:
            logger.error(f"Error getting window title: {e}")
            logger.error(traceback.format_exc())
            return ""
    
    def get_window_class(self, window) -> str:
        """
        Get the class name of a window.
        
        Args:
            window: Window object
            
        Returns:
            Window class name
        """
        try:
            if hasattr(window, 'class_name'):
                if callable(window.class_name):
                    class_name = window.class_name()
                else:
                    class_name = window.class_name
                logger.info(f"Window class name: {class_name}")
                return class_name
            else:
                logger.warning("Window doesn't have class_name method")
                return ""
        except Exception as e:
            logger.error(f"Error getting window class name: {e}")
            logger.error(traceback.format_exc())
            return ""
    
    def get_window_handle(self, window) -> Optional[int]:
        """
        Get the handle of a window.
        
        Args:
            window: Window object
            
        Returns:
            Window handle or None if failed
        """
        try:
            if hasattr(window, 'handle'):
                handle = window.handle
                logger.info(f"Window handle: {handle}")
                return handle
            else:
                logger.warning("Window doesn't have handle attribute")
                return None
        except Exception as e:
            logger.error(f"Error getting window handle: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_pid(self, window) -> Optional[int]:
        """
        Get the process ID of a window.
        
        Args:
            window: Window object
            
        Returns:
            Window process ID or None if failed
        """
        try:
            if hasattr(window, 'process_id'):
                if callable(window.process_id):
                    pid = window.process_id()
                else:
                    pid = window.process_id
                logger.info(f"Window process ID: {pid}")
                return pid
            else:
                logger.warning("Window doesn't have process_id method")
                return None
        except Exception as e:
            logger.error(f"Error getting window process ID: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_rectangle(self, window) -> Optional[Any]:
        """
        Get the rectangle (position and size) of a window.
        
        Args:
            window: Window object
            
        Returns:
            Window rectangle or None if failed
        """
        try:
            if hasattr(window, 'rectangle'):
                rect = window.rectangle()
                logger.info(f"Window rectangle: {rect}")
                return rect
            else:
                logger.warning("Window doesn't have rectangle method")
                return None
        except Exception as e:
            logger.error(f"Error getting window rectangle: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def set_window_position(self, window, x, y) -> bool:
        """
        Set the position of a window.
        
        Args:
            window: Window object
            x: X coordinate
            y: Y coordinate
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Setting window position to ({x}, {y})")
            
            if hasattr(window, 'move_window'):
                window.move_window(x, y)
                logger.info(f"Set window position to ({x}, {y})")
                return True
            else:
                logger.warning("Window doesn't have move_window method")
                return False
        except Exception as e:
            logger.error(f"Error setting window position: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def set_window_size(self, window, width, height) -> bool:
        """
        Set the size of a window.
        
        Args:
            window: Window object
            width: Width
            height: Height
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Setting window size to {width}x{height}")
            
            if hasattr(window, 'resize'):
                window.resize(width, height)
                logger.info(f"Set window size to {width}x{height}")
                return True
            else:
                logger.warning("Window doesn't have resize method")
                return False
        except Exception as e:
            logger.error(f"Error setting window size: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def set_window_position_and_size(self, window, x, y, width, height) -> bool:
        """
        Set the position and size of a window.
        
        Args:
            window: Window object
            x: X coordinate
            y: Y coordinate
            width: Width
            height: Height
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Setting window position and size to ({x}, {y}) {width}x{height}")
            
            if hasattr(window, 'move_window'):
                window.move_window(x, y, width, height)
                logger.info(f"Set window position and size to ({x}, {y}) {width}x{height}")
                return True
            else:
                # Try to set position and size separately
                position_success = self.set_window_position(window, x, y)
                size_success = self.set_window_size(window, width, height)
                return position_success and size_success
        except Exception as e:
            logger.error(f"Error setting window position and size: {e}")
            logger.error(traceback.format_exc())
            return False
    def get_window_state(self, window) -> str:
        """
        Get the state of a window (maximized, minimized, normal).
        
        Args:
            window: Window object
            
        Returns:
            Window state ("maximized", "minimized", "normal", or "unknown")
        """
        try:
            if hasattr(window, 'is_maximized') and window.is_maximized():
                logger.info("Window state: maximized")
                return "maximized"
            elif hasattr(window, 'is_minimized') and window.is_minimized():
                logger.info("Window state: minimized")
                return "minimized"
            else:
                logger.info("Window state: normal")
                return "normal"
        except Exception as e:
            logger.error(f"Error getting window state: {e}")
            logger.error(traceback.format_exc())
            return "unknown"
    
    def is_window_maximized(self, window) -> bool:
        """
        Check if a window is maximized.
        
        Args:
            window: Window object
            
        Returns:
            True if maximized, False otherwise
        """
        try:
            if hasattr(window, 'is_maximized'):
                return window.is_maximized()
            else:
                logger.warning("Window doesn't have is_maximized method")
                return False
        except Exception as e:
            logger.error(f"Error checking if window is maximized: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_minimized(self, window) -> bool:
        """
        Check if a window is minimized.
        
        Args:
            window: Window object
            
        Returns:
            True if minimized, False otherwise
        """
        try:
            if hasattr(window, 'is_minimized'):
                return window.is_minimized()
            else:
                logger.warning("Window doesn't have is_minimized method")
                return False
        except Exception as e:
            logger.error(f"Error checking if window is minimized: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_normal(self, window) -> bool:
        """
        Check if a window is in normal state (not maximized or minimized).
        
        Args:
            window: Window object
            
        Returns:
            True if normal, False otherwise
        """
        try:
            return not self.is_window_maximized(window) and not self.is_window_minimized(window)
        except Exception as e:
            logger.error(f"Error checking if window is normal: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_topmost(self, window) -> bool:
        """
        Check if a window is topmost (always on top).
        
        Args:
            window: Window object
            
        Returns:
            True if topmost, False otherwise
        """
        try:
            if hasattr(window, 'is_topmost'):
                return window.is_topmost()
            else:
                logger.warning("Window doesn't have is_topmost method")
                return False
        except Exception as e:
            logger.error(f"Error checking if window is topmost: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_enabled(self, window) -> bool:
        """
        Check if a window is enabled.
        
        Args:
            window: Window object
            
        Returns:
            True if enabled, False otherwise
        """
        try:
            if hasattr(window, 'is_enabled'):
                return window.is_enabled()
            else:
                logger.warning("Window doesn't have is_enabled method")
                return False
        except Exception as e:
            logger.error(f"Error checking if window is enabled: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_visible(self, window) -> bool:
        """
        Check if a window is visible.
        
        Args:
            window: Window object
            
        Returns:
            True if visible, False otherwise
        """
        try:
            if hasattr(window, 'is_visible'):
                return window.is_visible()
            else:
                logger.warning("Window doesn't have is_visible method")
                return False
        except Exception as e:
            logger.error(f"Error checking if window is visible: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_active(self, window) -> bool:
        """
        Check if a window is active (in foreground).
        
        Args:
            window: Window object
            
        Returns:
            True if active, False otherwise
        """
        try:
            if hasattr(window, 'is_active'):
                return window.is_active()
            else:
                logger.warning("Window doesn't have is_active method")
                return False
        except Exception as e:
            logger.error(f"Error checking if window is active: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_focused(self, window) -> bool:
        """
        Check if a window has focus.
        
        Args:
            window: Window object
            
        Returns:
            True if focused, False otherwise
        """
        try:
            if hasattr(window, 'has_focus'):
                return window.has_focus()
            else:
                # Fallback to is_active
                return self.is_window_active(window)
        except Exception as e:
            logger.error(f"Error checking if window is focused: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_modal(self, window) -> bool:
        """
        Check if a window is modal.
        
        Args:
            window: Window object
            
        Returns:
            True if modal, False otherwise
        """
        try:
            if hasattr(window, 'is_dialog'):
                return window.is_dialog()
            else:
                logger.warning("Window doesn't have is_dialog method")
                return False
        except Exception as e:
            logger.error(f"Error checking if window is modal: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_child(self, window, parent_window) -> bool:
        """
        Check if a window is a child of another window.
        
        Args:
            window: Window object
            parent_window: Parent window object
            
        Returns:
            True if child, False otherwise
        """
        try:
            # Get the parent of the window
            if hasattr(window, 'parent'):
                parent = window.parent()
                if parent:
                    # Check if the parent is the same as parent_window
                    if hasattr(parent, 'handle') and hasattr(parent_window, 'handle'):
                        return parent.handle == parent_window.handle
            
            logger.warning("Could not determine if window is a child")
            return False
        except Exception as e:
            logger.error(f"Error checking if window is a child: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_same(self, window1, window2) -> bool:
        """
        Check if two window objects refer to the same window.
        
        Args:
            window1: First window object
            window2: Second window object
            
        Returns:
            True if same, False otherwise
        """
        try:
            # Check if the handles are the same
            if hasattr(window1, 'handle') and hasattr(window2, 'handle'):
                return window1.handle == window2.handle
            
            # If no handles, check if the titles and class names are the same
            if (hasattr(window1, 'window_text') and hasattr(window2, 'window_text') and
                hasattr(window1, 'class_name') and hasattr(window2, 'class_name')):
                return (window1.window_text() == window2.window_text() and
                        window1.class_name() == window2.class_name())
            
            logger.warning("Could not determine if windows are the same")
            return False
        except Exception as e:
            logger.error(f"Error checking if windows are the same: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_window_process_name(self, window) -> Optional[str]:
        """
        Get the process name of a window.
        
        Args:
            window: Window object
            
        Returns:
            Process name or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process name
            import psutil
            process = psutil.Process(pid)
            process_name = process.name()
            
            logger.info(f"Window process name: {process_name}")
            return process_name
        except ImportError:
            logger.warning("psutil not available, cannot get process name")
            return None
        except Exception as e:
            logger.error(f"Error getting window process name: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_path(self, window) -> Optional[str]:
        """
        Get the process path of a window.
        
        Args:
            window: Window object
            
        Returns:
            Process path or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process path
            import psutil
            process = psutil.Process(pid)
            process_path = process.exe()
            
            logger.info(f"Window process path: {process_path}")
            return process_path
        except ImportError:
            logger.warning("psutil not available, cannot get process path")
            return None
        except Exception as e:
            logger.error(f"Error getting window process path: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_command_line(self, window) -> Optional[str]:
        """
        Get the process command line of a window.
        
        Args:
            window: Window object
            
        Returns:
            Process command line or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process command line
            import psutil
            process = psutil.Process(pid)
            cmdline = process.cmdline()
            
            # Join the command line arguments
            command_line = ' '.join(cmdline)
            
            logger.info(f"Window process command line: {command_line}")
            return command_line
        except ImportError:
            logger.warning("psutil not available, cannot get process command line")
            return None
        except Exception as e:
            logger.error(f"Error getting window process command line: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_creation_time(self, window) -> Optional[float]:
        """
        Get the process creation time of a window.
        
        Args:
            window: Window object
            
        Returns:
            Process creation time (timestamp) or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process creation time
            import psutil
            process = psutil.Process(pid)
            creation_time = process.create_time()
            
            logger.info(f"Window process creation time: {creation_time}")
            return creation_time
        except ImportError:
            logger.warning("psutil not available, cannot get process creation time")
            return None
        except Exception as e:
            logger.error(f"Error getting window process creation time: {e}")
            logger.error(traceback.format_exc())
            return None
    def get_window_process_username(self, window) -> Optional[str]:
        """
        Get the username of the process owner of a window.
        
        Args:
            window: Window object
            
        Returns:
            Username or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process username
            import psutil
            process = psutil.Process(pid)
            username = process.username()
            
            logger.info(f"Window process username: {username}")
            return username
        except ImportError:
            logger.warning("psutil not available, cannot get process username")
            return None
        except Exception as e:
            logger.error(f"Error getting window process username: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_status(self, window) -> Optional[str]:
        """
        Get the status of the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            Process status or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process status
            import psutil
            process = psutil.Process(pid)
            status = process.status()
            
            logger.info(f"Window process status: {status}")
            return status
        except ImportError:
            logger.warning("psutil not available, cannot get process status")
            return None
        except Exception as e:
            logger.error(f"Error getting window process status: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_memory_info(self, window) -> Optional[Dict[str, int]]:
        """
        Get memory information of the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            Dictionary with memory information or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process memory info
            import psutil
            process = psutil.Process(pid)
            memory_info = process.memory_info()
            
            # Convert to dictionary
            memory_dict = {
                'rss': memory_info.rss,  # Resident Set Size
                'vms': memory_info.vms,  # Virtual Memory Size
                'shared': getattr(memory_info, 'shared', 0),  # Shared memory
                'text': getattr(memory_info, 'text', 0),  # Text (code)
                'data': getattr(memory_info, 'data', 0),  # Data + stack
                'lib': getattr(memory_info, 'lib', 0),  # Library
                'dirty': getattr(memory_info, 'dirty', 0)  # Dirty pages
            }
            
            logger.info(f"Window process memory info: {memory_dict}")
            return memory_dict
        except ImportError:
            logger.warning("psutil not available, cannot get process memory info")
            return None
        except Exception as e:
            logger.error(f"Error getting window process memory info: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_cpu_percent(self, window) -> Optional[float]:
        """
        Get CPU usage percentage of the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            CPU usage percentage or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process CPU usage
            import psutil
            process = psutil.Process(pid)
            cpu_percent = process.cpu_percent(interval=0.1)
            
            logger.info(f"Window process CPU usage: {cpu_percent}%")
            return cpu_percent
        except ImportError:
            logger.warning("psutil not available, cannot get process CPU usage")
            return None
        except Exception as e:
            logger.error(f"Error getting window process CPU usage: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_threads(self, window) -> Optional[List[Dict[str, Any]]]:
        """
        Get threads of the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            List of thread dictionaries or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process threads
            import psutil
            process = psutil.Process(pid)
            threads = process.threads()
            
            # Convert to list of dictionaries
            thread_list = []
            for thread in threads:
                thread_dict = {
                    'id': thread.id,
                    'user_time': thread.user_time,
                    'system_time': thread.system_time
                }
                thread_list.append(thread_dict)
            
            logger.info(f"Window process has {len(thread_list)} threads")
            return thread_list
        except ImportError:
            logger.warning("psutil not available, cannot get process threads")
            return None
        except Exception as e:
            logger.error(f"Error getting window process threads: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_open_files(self, window) -> Optional[List[str]]:
        """
        Get open files of the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            List of file paths or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process open files
            import psutil
            process = psutil.Process(pid)
            open_files = process.open_files()
            
            # Extract file paths
            file_paths = [file.path for file in open_files]
            
            logger.info(f"Window process has {len(file_paths)} open files")
            return file_paths
        except ImportError:
            logger.warning("psutil not available, cannot get process open files")
            return None
        except Exception as e:
            logger.error(f"Error getting window process open files: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_connections(self, window) -> Optional[List[Dict[str, Any]]]:
        """
        Get network connections of the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            List of connection dictionaries or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process connections
            import psutil
            process = psutil.Process(pid)
            connections = process.connections()
            
            # Convert to list of dictionaries
            connection_list = []
            for conn in connections:
                conn_dict = {
                    'fd': conn.fd,
                    'family': conn.family,
                    'type': conn.type,
                    'laddr': conn.laddr._asdict() if conn.laddr else None,
                    'raddr': conn.raddr._asdict() if conn.raddr else None,
                    'status': conn.status
                }
                connection_list.append(conn_dict)
            
            logger.info(f"Window process has {len(connection_list)} network connections")
            return connection_list
        except ImportError:
            logger.warning("psutil not available, cannot get process connections")
            return None
        except Exception as e:
            logger.error(f"Error getting window process connections: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_environ(self, window) -> Optional[Dict[str, str]]:
        """
        Get environment variables of the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            Dictionary of environment variables or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process environment variables
            import psutil
            process = psutil.Process(pid)
            environ = process.environ()
            
            logger.info(f"Window process has {len(environ)} environment variables")
            return environ
        except ImportError:
            logger.warning("psutil not available, cannot get process environment variables")
            return None
        except Exception as e:
            logger.error(f"Error getting window process environment variables: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_children(self, window) -> Optional[List[int]]:
        """
        Get child processes of the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            List of child process IDs or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process children
            import psutil
            process = psutil.Process(pid)
            children = process.children()
            
            # Extract PIDs
            child_pids = [child.pid for child in children]
            
            logger.info(f"Window process has {len(child_pids)} child processes")
            return child_pids
        except ImportError:
            logger.warning("psutil not available, cannot get process children")
            return None
        except Exception as e:
            logger.error(f"Error getting window process children: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_parent(self, window) -> Optional[int]:
        """
        Get parent process of the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            Parent process ID or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process parent
            import psutil
            process = psutil.Process(pid)
            parent = process.parent()
            
            if parent:
                parent_pid = parent.pid
                logger.info(f"Window process parent PID: {parent_pid}")
                return parent_pid
            else:
                logger.info("Window process has no parent")
                return None
        except ImportError:
            logger.warning("psutil not available, cannot get process parent")
            return None
        except Exception as e:
            logger.error(f"Error getting window process parent: {e}")
            logger.error(traceback.format_exc())
            return None
    def get_window_process_nice(self, window) -> Optional[int]:
        """
        Get the nice value (priority) of the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            Nice value or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process nice value
            import psutil
            process = psutil.Process(pid)
            nice = process.nice()
            
            logger.info(f"Window process nice value: {nice}")
            return nice
        except ImportError:
            logger.warning("psutil not available, cannot get process nice value")
            return None
        except Exception as e:
            logger.error(f"Error getting window process nice value: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def set_window_process_nice(self, window, nice_value) -> bool:
        """
        Set the nice value (priority) of the process of a window.
        
        Args:
            window: Window object
            nice_value: Nice value to set
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return False
            
            # Set the process nice value
            import psutil
            process = psutil.Process(pid)
            process.nice(nice_value)
            
            logger.info(f"Set window process nice value to {nice_value}")
            return True
        except ImportError:
            logger.warning("psutil not available, cannot set process nice value")
            return False
        except Exception as e:
            logger.error(f"Error setting window process nice value: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_window_process_io_counters(self, window) -> Optional[Dict[str, int]]:
        """
        Get I/O counters of the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            Dictionary with I/O counters or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process I/O counters
            import psutil
            process = psutil.Process(pid)
            io_counters = process.io_counters()
            
            # Convert to dictionary
            io_dict = {
                'read_count': io_counters.read_count,
                'write_count': io_counters.write_count,
                'read_bytes': io_counters.read_bytes,
                'write_bytes': io_counters.write_bytes,
                'other_count': getattr(io_counters, 'other_count', 0),
                'other_bytes': getattr(io_counters, 'other_bytes', 0)
            }
            
            logger.info(f"Window process I/O counters: {io_dict}")
            return io_dict
        except ImportError:
            logger.warning("psutil not available, cannot get process I/O counters")
            return None
        except Exception as e:
            logger.error(f"Error getting window process I/O counters: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_num_handles(self, window) -> Optional[int]:
        """
        Get the number of handles used by the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            Number of handles or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process number of handles
            import psutil
            process = psutil.Process(pid)
            num_handles = process.num_handles()
            
            logger.info(f"Window process number of handles: {num_handles}")
            return num_handles
        except ImportError:
            logger.warning("psutil not available, cannot get process number of handles")
            return None
        except AttributeError:
            logger.warning("num_handles method not available (non-Windows platform)")
            return None
        except Exception as e:
            logger.error(f"Error getting window process number of handles: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_num_ctx_switches(self, window) -> Optional[Dict[str, int]]:
        """
        Get the number of context switches of the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            Dictionary with context switch counts or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process number of context switches
            import psutil
            process = psutil.Process(pid)
            ctx_switches = process.num_ctx_switches()
            
            # Convert to dictionary
            ctx_dict = {
                'voluntary': ctx_switches.voluntary,
                'involuntary': ctx_switches.involuntary
            }
            
            logger.info(f"Window process context switches: {ctx_dict}")
            return ctx_dict
        except ImportError:
            logger.warning("psutil not available, cannot get process context switches")
            return None
        except Exception as e:
            logger.error(f"Error getting window process context switches: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_num_fds(self, window) -> Optional[int]:
        """
        Get the number of file descriptors used by the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            Number of file descriptors or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process number of file descriptors
            import psutil
            process = psutil.Process(pid)
            num_fds = process.num_fds()
            
            logger.info(f"Window process number of file descriptors: {num_fds}")
            return num_fds
        except ImportError:
            logger.warning("psutil not available, cannot get process number of file descriptors")
            return None
        except AttributeError:
            logger.warning("num_fds method not available (Windows platform)")
            return None
        except Exception as e:
            logger.error(f"Error getting window process number of file descriptors: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_window_process_cpu_affinity(self, window) -> Optional[List[int]]:
        """
        Get the CPU affinity of the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            List of CPU indices or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process CPU affinity
            import psutil
            process = psutil.Process(pid)
            cpu_affinity = process.cpu_affinity()
            
            logger.info(f"Window process CPU affinity: {cpu_affinity}")
            return cpu_affinity
        except ImportError:
            logger.warning("psutil not available, cannot get process CPU affinity")
            return None
        except AttributeError:
            logger.warning("cpu_affinity method not available (non-Windows/Linux platform)")
            return None
        except Exception as e:
            logger.error(f"Error getting window process CPU affinity: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def set_window_process_cpu_affinity(self, window, cpu_list) -> bool:
        """
        Set the CPU affinity of the process of a window.
        
        Args:
            window: Window object
            cpu_list: List of CPU indices
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return False
            
            # Set the process CPU affinity
            import psutil
            process = psutil.Process(pid)
            process.cpu_affinity(cpu_list)
            
            logger.info(f"Set window process CPU affinity to {cpu_list}")
            return True
        except ImportError:
            logger.warning("psutil not available, cannot set process CPU affinity")
            return False
        except AttributeError:
            logger.warning("cpu_affinity method not available (non-Windows/Linux platform)")
            return False
        except Exception as e:
            logger.error(f"Error setting window process CPU affinity: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_window_process_memory_maps(self, window) -> Optional[List[Dict[str, Any]]]:
        """
        Get memory maps of the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            List of memory map dictionaries or None if failed
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return None
            
            # Get the process memory maps
            import psutil
            process = psutil.Process(pid)
            memory_maps = process.memory_maps()
            
            # Convert to list of dictionaries
            map_list = []
            for mmap in memory_maps:
                map_dict = mmap._asdict()
                map_list.append(map_dict)
            
            logger.info(f"Window process has {len(map_list)} memory maps")
            return map_list
        except ImportError:
            logger.warning("psutil not available, cannot get process memory maps")
            return None
        except Exception as e:
            logger.error(f"Error getting window process memory maps: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def suspend_window_process(self, window) -> bool:
        """
        Suspend the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return False
            
            # Suspend the process
            import psutil
            process = psutil.Process(pid)
            process.suspend()
            
            logger.info(f"Suspended window process with PID {pid}")
            return True
        except ImportError:
            logger.warning("psutil not available, cannot suspend process")
            return False
        except AttributeError:
            logger.warning("suspend method not available (non-Windows platform)")
            return False
        except Exception as e:
            logger.error(f"Error suspending window process: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def resume_window_process(self, window) -> bool:
        """
        Resume the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return False
            
            # Resume the process
            import psutil
            process = psutil.Process(pid)
            process.resume()
            
            logger.info(f"Resumed window process with PID {pid}")
            return True
        except ImportError:
            logger.warning("psutil not available, cannot resume process")
            return False
        except AttributeError:
            logger.warning("resume method not available (non-Windows platform)")
            return False
        except Exception as e:
            logger.error(f"Error resuming window process: {e}")
            logger.error(traceback.format_exc())
            return False
    def terminate_window_process(self, window) -> bool:
        """
        Terminate the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return False
            
            # Terminate the process
            import psutil
            process = psutil.Process(pid)
            process.terminate()
            
            logger.info(f"Terminated window process with PID {pid}")
            return True
        except ImportError:
            logger.warning("psutil not available, cannot terminate process")
            return False
        except Exception as e:
            logger.error(f"Error terminating window process: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def kill_window_process(self, window) -> bool:
        """
        Kill the process of a window.
        
        Args:
            window: Window object
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return False
            
            # Kill the process
            import psutil
            process = psutil.Process(pid)
            process.kill()
            
            logger.info(f"Killed window process with PID {pid}")
            return True
        except ImportError:
            logger.warning("psutil not available, cannot kill process")
            return False
        except Exception as e:
            logger.error(f"Error killing window process: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def wait_for_window_process_exit(self, window, timeout=None) -> bool:
        """
        Wait for the process of a window to exit.
        
        Args:
            window: Window object
            timeout: Timeout in seconds
            
        Returns:
            True if process exited, False otherwise
        """
        timeout = timeout or self.default_timeout
        
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return False
            
            # Wait for the process to exit
            import psutil
            process = psutil.Process(pid)
            process.wait(timeout=timeout)
            
            logger.info(f"Window process with PID {pid} exited")
            return True
        except ImportError:
            logger.warning("psutil not available, cannot wait for process exit")
            return False
        except psutil.TimeoutExpired:
            logger.warning(f"Process did not exit within {timeout} seconds")
            return False
        except Exception as e:
            logger.error(f"Error waiting for window process exit: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_window_process_running(self, window) -> bool:
        """
        Check if the process of a window is running.
        
        Args:
            window: Window object
            
        Returns:
            True if running, False otherwise
        """
        try:
            # Get the process ID
            pid = self.get_window_pid(window)
            if pid is None:
                logger.warning("Could not get window process ID")
                return False
            
            # Check if the process is running
            import psutil
            return psutil.pid_exists(pid)
        except ImportError:
            logger.warning("psutil not available, using alternative method")
            
            # Alternative method using pywinauto
            try:
                return window.exists()
            except Exception:
                return False
        except Exception as e:
            logger.error(f"Error checking if window process is running: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_all_processes(self) -> List[Dict[str, Any]]:
        """
        Get information about all running processes.
        
        Returns:
            List of process dictionaries
        """
        try:
            import psutil
            
            processes = []
            for proc in psutil.process_iter(['pid', 'name', 'username', 'status']):
                try:
                    proc_info = proc.info
                    processes.append(proc_info)
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
            
            logger.info(f"Found {len(processes)} running processes")
            return processes
        except ImportError:
            logger.warning("psutil not available, cannot get all processes")
            return []
        except Exception as e:
            logger.error(f"Error getting all processes: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def find_process_by_name(self, name) -> List[int]:
        """
        Find processes by name.
        
        Args:
            name: Process name
            
        Returns:
            List of process IDs
        """
        try:
            import psutil
            
            pids = []
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if proc.info['name'].lower() == name.lower():
                        pids.append(proc.info['pid'])
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
            
            logger.info(f"Found {len(pids)} processes with name '{name}'")
            return pids
        except ImportError:
            logger.warning("psutil not available, cannot find processes by name")
            return []
        except Exception as e:
            logger.error(f"Error finding processes by name: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def find_process_by_path(self, path) -> List[int]:
        """
        Find processes by executable path.
        
        Args:
            path: Executable path
            
        Returns:
            List of process IDs
        """
        try:
            import psutil
            
            pids = []
            for proc in psutil.process_iter(['pid', 'exe']):
                try:
                    if proc.info['exe'] and proc.info['exe'].lower() == path.lower():
                        pids.append(proc.info['pid'])
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
            
            logger.info(f"Found {len(pids)} processes with path '{path}'")
            return pids
        except ImportError:
            logger.warning("psutil not available, cannot find processes by path")
            return []
        except Exception as e:
            logger.error(f"Error finding processes by path: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def find_process_by_cmdline(self, cmdline_pattern) -> List[int]:
        """
        Find processes by command line pattern.
        
        Args:
            cmdline_pattern: Command line pattern to match
            
        Returns:
            List of process IDs
        """
        try:
            import psutil
            import re
            
            pattern = re.compile(cmdline_pattern, re.IGNORECASE)
            
            pids = []
            for proc in psutil.process_iter(['pid', 'cmdline']):
                try:
                    cmdline = ' '.join(proc.info['cmdline'] or [])
                    if pattern.search(cmdline):
                        pids.append(proc.info['pid'])
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
            
            logger.info(f"Found {len(pids)} processes matching cmdline pattern '{cmdline_pattern}'")
            return pids
        except ImportError:
            logger.warning("psutil not available, cannot find processes by cmdline")
            return []
        except Exception as e:
            logger.error(f"Error finding processes by cmdline: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def find_process_by_window_title(self, title_pattern) -> List[int]:
        """
        Find processes by window title pattern.
        
        Args:
            title_pattern: Window title pattern to match
            
        Returns:
            List of process IDs
        """
        try:
            import re
            
            pattern = re.compile(title_pattern, re.IGNORECASE)
            
            # Get all windows
            windows = self.desktop.windows()
            
            pids = set()
            for window in windows:
                try:
                    title = window.window_text()
                    if pattern.search(title):
                        pid = window.process_id()
                        pids.add(pid)
                except Exception:
                    pass
            
            logger.info(f"Found {len(pids)} processes with windows matching title pattern '{title_pattern}'")
            return list(pids)
        except Exception as e:
            logger.error(f"Error finding processes by window title: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def find_process_by_window_class(self, class_pattern) -> List[int]:
        """
        Find processes by window class pattern.
        
        Args:
            class_pattern: Window class pattern to match
            
        Returns:
            List of process IDs
        """
        try:
            import re
            
            pattern = re.compile(class_pattern, re.IGNORECASE)
            
            # Get all windows
            windows = self.desktop.windows()
            
            pids = set()
            for window in windows:
                try:
                    class_name = window.class_name()
                    if pattern.search(class_name):
                        pid = window.process_id()
                        pids.add(pid)
                except Exception:
                    pass
            
            logger.info(f"Found {len(pids)} processes with windows matching class pattern '{class_pattern}'")
            return list(pids)
        except Exception as e:
            logger.error(f"Error finding processes by window class: {e}")
            logger.error(traceback.format_exc())
            return []
    def get_system_info(self) -> Dict[str, Any]:
        """
        Get system information.
        
        Returns:
            Dictionary with system information
        """
        try:
            import platform
            import psutil
            
            # Get basic system information
            system_info = {
                'system': platform.system(),
                'node': platform.node(),
                'release': platform.release(),
                'version': platform.version(),
                'machine': platform.machine(),
                'processor': platform.processor(),
                'python_version': platform.python_version(),
                'cpu_count': psutil.cpu_count(logical=False),
                'cpu_count_logical': psutil.cpu_count(logical=True),
                'memory_total': psutil.virtual_memory().total,
                'memory_available': psutil.virtual_memory().available,
                'disk_usage': {path: psutil.disk_usage(path)._asdict() for path in psutil.disk_partitions() if psutil.disk_usage(path).total > 0},
                'boot_time': psutil.boot_time()
            }
            
            logger.info("Got system information")
            return system_info
        except ImportError:
            logger.warning("psutil not available, getting limited system information")
            
            # Get limited system information without psutil
            import platform
            
            system_info = {
                'system': platform.system(),
                'node': platform.node(),
                'release': platform.release(),
                'version': platform.version(),
                'machine': platform.machine(),
                'processor': platform.processor(),
                'python_version': platform.python_version()
            }
            
            logger.info("Got limited system information")
            return system_info
        except Exception as e:
            logger.error(f"Error getting system information: {e}")
            logger.error(traceback.format_exc())
            return {}
    
    def get_system_cpu_usage(self) -> float:
        """
        Get system CPU usage.
        
        Returns:
            CPU usage percentage
        """
        try:
            import psutil
            
            cpu_percent = psutil.cpu_percent(interval=0.1)
            logger.debug(f"System CPU usage: {cpu_percent}%")
            return cpu_percent
        except ImportError:
            logger.warning("psutil not available, cannot get system CPU usage")
            return 0.0
        except Exception as e:
            logger.error(f"Error getting system CPU usage: {e}")
            logger.error(traceback.format_exc())
            return 0.0
    
    def get_system_memory_usage(self) -> Dict[str, int]:
        """
        Get system memory usage.
        
        Returns:
            Dictionary with memory usage information
        """
        try:
            import psutil
            
            memory = psutil.virtual_memory()
            memory_dict = {
                'total': memory.total,
                'available': memory.available,
                'used': memory.used,
                'free': memory.free,
                'percent': memory.percent
            }
            
            logger.debug(f"System memory usage: {memory_dict['percent']}%")
            return memory_dict
        except ImportError:
            logger.warning("psutil not available, cannot get system memory usage")
            return {}
        except Exception as e:
            logger.error(f"Error getting system memory usage: {e}")
            logger.error(traceback.format_exc())
            return {}
    
    def get_system_disk_usage(self, path='/') -> Dict[str, int]:
        """
        Get system disk usage.
        
        Args:
            path: Disk path
            
        Returns:
            Dictionary with disk usage information
        """
        try:
            import psutil
            
            disk = psutil.disk_usage(path)
            disk_dict = {
                'total': disk.total,
                'used': disk.used,
                'free': disk.free,
                'percent': disk.percent
            }
            
            logger.debug(f"System disk usage for {path}: {disk_dict['percent']}%")
            return disk_dict
        except ImportError:
            logger.warning("psutil not available, cannot get system disk usage")
            return {}
        except Exception as e:
            logger.error(f"Error getting system disk usage: {e}")
            logger.error(traceback.format_exc())
            return {}
    
    def get_system_network_usage(self) -> Dict[str, Dict[str, int]]:
        """
        Get system network usage.
        
        Returns:
            Dictionary with network usage information
        """
        try:
            import psutil
            
            network = psutil.net_io_counters(pernic=True)
            network_dict = {nic: {
                'bytes_sent': stats.bytes_sent,
                'bytes_recv': stats.bytes_recv,
                'packets_sent': stats.packets_sent,
                'packets_recv': stats.packets_recv,
                'errin': stats.errin,
                'errout': stats.errout,
                'dropin': stats.dropin,
                'dropout': stats.dropout
            } for nic, stats in network.items()}
            
            logger.debug(f"Got system network usage for {len(network_dict)} interfaces")
            return network_dict
        except ImportError:
            logger.warning("psutil not available, cannot get system network usage")
            return {}
        except Exception as e:
            logger.error(f"Error getting system network usage: {e}")
            logger.error(traceback.format_exc())
            return {}
    
    def get_system_users(self) -> List[Dict[str, Any]]:
        """
        Get system users.
        
        Returns:
            List of user dictionaries
        """
        try:
            import psutil
            
            users = psutil.users()
            user_list = [{
                'name': user.name,
                'terminal': user.terminal,
                'host': user.host,
                'started': user.started,
                'pid': getattr(user, 'pid', None)
            } for user in users]
            
            logger.debug(f"Got {len(user_list)} system users")
            return user_list
        except ImportError:
            logger.warning("psutil not available, cannot get system users")
            return []
        except Exception as e:
            logger.error(f"Error getting system users: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_system_boot_time(self) -> float:
        """
        Get system boot time.
        
        Returns:
            Boot time timestamp
        """
        try:
            import psutil
            
            boot_time = psutil.boot_time()
            logger.debug(f"System boot time: {boot_time}")
            return boot_time
        except ImportError:
            logger.warning("psutil not available, cannot get system boot time")
            return 0.0
        except Exception as e:
            logger.error(f"Error getting system boot time: {e}")
            logger.error(traceback.format_exc())
            return 0.0
    
    def get_system_uptime(self) -> float:
        """
        Get system uptime.
        
        Returns:
            Uptime in seconds
        """
        try:
            import psutil
            import time
            
            boot_time = psutil.boot_time()
            uptime = time.time() - boot_time
            logger.debug(f"System uptime: {uptime} seconds")
            return uptime
        except ImportError:
            logger.warning("psutil not available, cannot get system uptime")
            return 0.0
        except Exception as e:
            logger.error(f"Error getting system uptime: {e}")
            logger.error(traceback.format_exc())
            return 0.0
    
    def get_system_load_average(self) -> List[float]:
        """
        Get system load average.
        
        Returns:
            List of load averages (1, 5, 15 minutes)
        """
        try:
            import psutil
            
            load_avg = psutil.getloadavg()
            logger.debug(f"System load average: {load_avg}")
            return list(load_avg)
        except ImportError:
            logger.warning("psutil not available, cannot get system load average")
            return [0.0, 0.0, 0.0]
        except Exception as e:
            logger.error(f"Error getting system load average: {e}")
            logger.error(traceback.format_exc())
            return [0.0, 0.0, 0.0]
    
    def get_system_cpu_times(self) -> Dict[str, float]:
        """
        Get system CPU times.
        
        Returns:
            Dictionary with CPU times
        """
        try:
            import psutil
            
            cpu_times = psutil.cpu_times()
            cpu_dict = {
                'user': cpu_times.user,
                'system': cpu_times.system,
                'idle': cpu_times.idle,
                'nice': getattr(cpu_times, 'nice', 0.0),
                'iowait': getattr(cpu_times, 'iowait', 0.0),
                'irq': getattr(cpu_times, 'irq', 0.0),
                'softirq': getattr(cpu_times, 'softirq', 0.0),
                'steal': getattr(cpu_times, 'steal', 0.0),
                'guest': getattr(cpu_times, 'guest', 0.0),
                'guest_nice': getattr(cpu_times, 'guest_nice', 0.0)
            }
            
            logger.debug(f"Got system CPU times")
            return cpu_dict
        except ImportError:
            logger.warning("psutil not available, cannot get system CPU times")
            return {}
        except Exception as e:
            logger.error(f"Error getting system CPU times: {e}")
            logger.error(traceback.format_exc())
            return {}
    
    def get_system_cpu_stats(self) -> Dict[str, int]:
        """
        Get system CPU statistics.
        
        Returns:
            Dictionary with CPU statistics
        """
        try:
            import psutil
            
            cpu_stats = psutil.cpu_stats()
            stats_dict = {
                'ctx_switches': cpu_stats.ctx_switches,
                'interrupts': cpu_stats.interrupts,
                'soft_interrupts': cpu_stats.soft_interrupts,
                'syscalls': getattr(cpu_stats, 'syscalls', 0)
            }
            
            logger.debug(f"Got system CPU statistics")
            return stats_dict
        except ImportError:
            logger.warning("psutil not available, cannot get system CPU statistics")
            return {}
        except Exception as e:
            logger.error(f"Error getting system CPU statistics: {e}")
            logger.error(traceback.format_exc())
            return {}
    def get_system_cpu_freq(self) -> Dict[str, float]:
        """
        Get system CPU frequency.
        
        Returns:
            Dictionary with CPU frequency information
        """
        try:
            import psutil
            
            cpu_freq = psutil.cpu_freq()
            if cpu_freq:
                freq_dict = {
                    'current': cpu_freq.current,
                    'min': cpu_freq.min,
                    'max': cpu_freq.max
                }
                
                logger.debug(f"System CPU frequency: {freq_dict['current']} MHz")
                return freq_dict
            else:
                logger.warning("CPU frequency information not available")
                return {}
        except ImportError:
            logger.warning("psutil not available, cannot get system CPU frequency")
            return {}
        except Exception as e:
            logger.error(f"Error getting system CPU frequency: {e}")
            logger.error(traceback.format_exc())
            return {}
    
    def get_system_swap_memory(self) -> Dict[str, int]:
        """
        Get system swap memory usage.
        
        Returns:
            Dictionary with swap memory usage information
        """
        try:
            import psutil
            
            swap = psutil.swap_memory()
            swap_dict = {
                'total': swap.total,
                'used': swap.used,
                'free': swap.free,
                'percent': swap.percent,
                'sin': swap.sin,
                'sout': swap.sout
            }
            
            logger.debug(f"System swap memory usage: {swap_dict['percent']}%")
            return swap_dict
        except ImportError:
            logger.warning("psutil not available, cannot get system swap memory usage")
            return {}
        except Exception as e:
            logger.error(f"Error getting system swap memory usage: {e}")
            logger.error(traceback.format_exc())
            return {}
    
    def get_system_disk_partitions(self) -> List[Dict[str, str]]:
        """
        Get system disk partitions.
        
        Returns:
            List of partition dictionaries
        """
        try:
            import psutil
            
            partitions = psutil.disk_partitions()
            partition_list = [{
                'device': partition.device,
                'mountpoint': partition.mountpoint,
                'fstype': partition.fstype,
                'opts': partition.opts
            } for partition in partitions]
            
            logger.debug(f"Got {len(partition_list)} system disk partitions")
            return partition_list
        except ImportError:
            logger.warning("psutil not available, cannot get system disk partitions")
            return []
        except Exception as e:
            logger.error(f"Error getting system disk partitions: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_system_disk_io_counters(self) -> Dict[str, Dict[str, int]]:
        """
        Get system disk I/O counters.
        
        Returns:
            Dictionary with disk I/O counters
        """
        try:
            import psutil
            
            disk_io = psutil.disk_io_counters(perdisk=True)
            disk_io_dict = {disk: {
                'read_count': stats.read_count,
                'write_count': stats.write_count,
                'read_bytes': stats.read_bytes,
                'write_bytes': stats.write_bytes,
                'read_time': stats.read_time,
                'write_time': stats.write_time
            } for disk, stats in disk_io.items()}
            
            logger.debug(f"Got system disk I/O counters for {len(disk_io_dict)} disks")
            return disk_io_dict
        except ImportError:
            logger.warning("psutil not available, cannot get system disk I/O counters")
            return {}
        except Exception as e:
            logger.error(f"Error getting system disk I/O counters: {e}")
            logger.error(traceback.format_exc())
            return {}
    
    def get_system_network_connections(self) -> List[Dict[str, Any]]:
        """
        Get system network connections.
        
        Returns:
            List of connection dictionaries
        """
        try:
            import psutil
            
            connections = psutil.net_connections()
            connection_list = [{
                'fd': conn.fd,
                'family': conn.family,
                'type': conn.type,
                'laddr': conn.laddr._asdict() if conn.laddr else None,
                'raddr': conn.raddr._asdict() if conn.raddr else None,
                'status': conn.status,
                'pid': conn.pid
            } for conn in connections]
            
            logger.debug(f"Got {len(connection_list)} system network connections")
            return connection_list
        except ImportError:
            logger.warning("psutil not available, cannot get system network connections")
            return []
        except Exception as e:
            logger.error(f"Error getting system network connections: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_system_network_addresses(self) -> Dict[str, List[Dict[str, str]]]:
        """
        Get system network addresses.
        
        Returns:
            Dictionary with network addresses
        """
        try:
            import psutil
            
            addresses = psutil.net_if_addrs()
            address_dict = {interface: [{
                'family': addr.family,
                'address': addr.address,
                'netmask': addr.netmask,
                'broadcast': addr.broadcast,
                'ptp': addr.ptp
            } for addr in addrs] for interface, addrs in addresses.items()}
            
            logger.debug(f"Got system network addresses for {len(address_dict)} interfaces")
            return address_dict
        except ImportError:
            logger.warning("psutil not available, cannot get system network addresses")
            return {}
        except Exception as e:
            logger.error(f"Error getting system network addresses: {e}")
            logger.error(traceback.format_exc())
            return {}
    
    def get_system_network_stats(self) -> Dict[str, Dict[str, bool]]:
        """
        Get system network statistics.
        
        Returns:
            Dictionary with network statistics
        """
        try:
            import psutil
            
            stats = psutil.net_if_stats()
            stats_dict = {interface: {
                'isup': stat.isup,
                'duplex': stat.duplex,
                'speed': stat.speed,
                'mtu': stat.mtu
            } for interface, stat in stats.items()}
            
            logger.debug(f"Got system network statistics for {len(stats_dict)} interfaces")
            return stats_dict
        except ImportError:
            logger.warning("psutil not available, cannot get system network statistics")
            return {}
        except Exception as e:
            logger.error(f"Error getting system network statistics: {e}")
            logger.error(traceback.format_exc())
            return {}
    
    def get_system_sensors_temperatures(self) -> Dict[str, List[Dict[str, Any]]]:
        """
        Get system temperature sensors.
        
        Returns:
            Dictionary with temperature sensors
        """
        try:
            import psutil
            
            temperatures = psutil.sensors_temperatures()
            temp_dict = {sensor: [{
                'label': temp.label,
                'current': temp.current,
                'high': temp.high,
                'critical': temp.critical
            } for temp in temps] for sensor, temps in temperatures.items()}
            
            logger.debug(f"Got system temperature sensors for {len(temp_dict)} sensors")
            return temp_dict
        except ImportError:
            logger.warning("psutil not available, cannot get system temperature sensors")
            return {}
        except AttributeError:
            logger.warning("sensors_temperatures method not available (non-Linux platform)")
            return {}
        except Exception as e:
            logger.error(f"Error getting system temperature sensors: {e}")
            logger.error(traceback.format_exc())
            return {}
    
    def get_system_sensors_fans(self) -> Dict[str, List[Dict[str, Any]]]:
        """
        Get system fan sensors.
        
        Returns:
            Dictionary with fan sensors
        """
        try:
            import psutil
            
            fans = psutil.sensors_fans()
            fan_dict = {sensor: [{
                'label': fan.label,
                'current': fan.current
            } for fan in fans_list] for sensor, fans_list in fans.items()}
            
            logger.debug(f"Got system fan sensors for {len(fan_dict)} sensors")
            return fan_dict
        except ImportError:
            logger.warning("psutil not available, cannot get system fan sensors")
            return {}
        except AttributeError:
            logger.warning("sensors_fans method not available (non-Linux platform)")
            return {}
        except Exception as e:
            logger.error(f"Error getting system fan sensors: {e}")
            logger.error(traceback.format_exc())
            return {}
    
    def get_system_sensors_battery(self) -> Dict[str, Any]:
        """
        Get system battery information.
        
        Returns:
            Dictionary with battery information
        """
        try:
            import psutil
            
            battery = psutil.sensors_battery()
            if battery:
                battery_dict = {
                    'percent': battery.percent,
                    'secsleft': battery.secsleft,
                    'power_plugged': battery.power_plugged
                }
                
                logger.debug(f"System battery: {battery_dict['percent']}%")
                return battery_dict
            else:
                logger.warning("No battery found")
                return {}
        except ImportError:
            logger.warning("psutil not available, cannot get system battery information")
            return {}
        except AttributeError:
            logger.warning("sensors_battery method not available")
            return {}
        except Exception as e:
            logger.error(f"Error getting system battery information: {e}")
            logger.error(traceback.format_exc())
            return {}
    def get_clipboard_text(self) -> Optional[str]:
        """
        Get text from the clipboard.
        
        Returns:
            Clipboard text or None if failed
        """
        try:
            import win32clipboard
            import win32con
            
            win32clipboard.OpenClipboard()
            try:
                if win32clipboard.IsClipboardFormatAvailable(win32con.CF_TEXT):
                    text = win32clipboard.GetClipboardData(win32con.CF_TEXT)
                    if isinstance(text, bytes):
                        text = text.decode('utf-8', errors='replace')
                    logger.info(f"Got text from clipboard: {text[:50]}...")
                    return text
                elif win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                    text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
                    logger.info(f"Got unicode text from clipboard: {text[:50]}...")
                    return text
                else:
                    logger.warning("No text format available in clipboard")
                    return None
            finally:
                win32clipboard.CloseClipboard()
        except ImportError:
            logger.warning("win32clipboard not available, cannot get clipboard text")
            return None
        except Exception as e:
            logger.error(f"Error getting clipboard text: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def set_clipboard_text(self, text: str) -> bool:
        """
        Set text to the clipboard.
        
        Args:
            text: Text to set
            
        Returns:
            True if successful, False otherwise
        """
        try:
            import win32clipboard
            import win32con
            
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardText(text, win32con.CF_UNICODETEXT)
                logger.info(f"Set text to clipboard: {text[:50]}...")
                return True
            finally:
                win32clipboard.CloseClipboard()
        except ImportError:
            logger.warning("win32clipboard not available, cannot set clipboard text")
            return False
        except Exception as e:
            logger.error(f"Error setting clipboard text: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def clear_clipboard(self) -> bool:
        """
        Clear the clipboard.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            import win32clipboard
            
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                logger.info("Cleared clipboard")
                return True
            finally:
                win32clipboard.CloseClipboard()
        except ImportError:
            logger.warning("win32clipboard not available, cannot clear clipboard")
            return False
        except Exception as e:
            logger.error(f"Error clearing clipboard: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_clipboard_formats(self) -> List[int]:
        """
        Get available clipboard formats.
        
        Returns:
            List of format identifiers
        """
        try:
            import win32clipboard
            
            win32clipboard.OpenClipboard()
            try:
                formats = []
                format_id = 0
                
                while True:
                    format_id = win32clipboard.EnumClipboardFormats(format_id)
                    if format_id == 0:
                        break
                    formats.append(format_id)
                
                logger.info(f"Got {len(formats)} clipboard formats")
                return formats
            finally:
                win32clipboard.CloseClipboard()
        except ImportError:
            logger.warning("win32clipboard not available, cannot get clipboard formats")
            return []
        except Exception as e:
            logger.error(f"Error getting clipboard formats: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_clipboard_format_name(self, format_id: int) -> Optional[str]:
        """
        Get the name of a clipboard format.
        
        Args:
            format_id: Format identifier
            
        Returns:
            Format name or None if failed
        """
        try:
            import win32clipboard
            
            name = win32clipboard.GetClipboardFormatName(format_id)
            logger.info(f"Clipboard format {format_id} name: {name}")
            return name
        except ImportError:
            logger.warning("win32clipboard not available, cannot get clipboard format name")
            return None
        except Exception as e:
            logger.error(f"Error getting clipboard format name: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def is_clipboard_format_available(self, format_id: int) -> bool:
        """
        Check if a clipboard format is available.
        
        Args:
            format_id: Format identifier
            
        Returns:
            True if available, False otherwise
        """
        try:
            import win32clipboard
            
            win32clipboard.OpenClipboard()
            try:
                available = win32clipboard.IsClipboardFormatAvailable(format_id)
                logger.info(f"Clipboard format {format_id} available: {available}")
                return available
            finally:
                win32clipboard.CloseClipboard()
        except ImportError:
            logger.warning("win32clipboard not available, cannot check clipboard format")
            return False
        except Exception as e:
            logger.error(f"Error checking clipboard format: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_clipboard_data(self, format_id: int) -> Optional[Any]:
        """
        Get data from the clipboard in a specific format.
        
        Args:
            format_id: Format identifier
            
        Returns:
            Clipboard data or None if failed
        """
        try:
            import win32clipboard
            
            win32clipboard.OpenClipboard()
            try:
                if win32clipboard.IsClipboardFormatAvailable(format_id):
                    data = win32clipboard.GetClipboardData(format_id)
                    logger.info(f"Got clipboard data for format {format_id}")
                    return data
                else:
                    logger.warning(f"Clipboard format {format_id} not available")
                    return None
            finally:
                win32clipboard.CloseClipboard()
        except ImportError:
            logger.warning("win32clipboard not available, cannot get clipboard data")
            return None
        except Exception as e:
            logger.error(f"Error getting clipboard data: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def set_clipboard_data(self, format_id: int, data: Any) -> bool:
        """
        Set data to the clipboard in a specific format.
        
        Args:
            format_id: Format identifier
            data: Data to set
            
        Returns:
            True if successful, False otherwise
        """
        try:
            import win32clipboard
            
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(format_id, data)
                logger.info(f"Set clipboard data for format {format_id}")
                return True
            finally:
                win32clipboard.CloseClipboard()
        except ImportError:
            logger.warning("win32clipboard not available, cannot set clipboard data")
            return False
        except Exception as e:
            logger.error(f"Error setting clipboard data: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def register_clipboard_format(self, format_name: str) -> Optional[int]:
        """
        Register a new clipboard format.
        
        Args:
            format_name: Format name
            
        Returns:
            Format identifier or None if failed
        """
        try:
            import win32clipboard
            
            format_id = win32clipboard.RegisterClipboardFormat(format_name)
            logger.info(f"Registered clipboard format '{format_name}' with ID {format_id}")
            return format_id
        except ImportError:
            logger.warning("win32clipboard not available, cannot register clipboard format")
            return None
        except Exception as e:
            logger.error(f"Error registering clipboard format: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_clipboard_sequence_number(self) -> Optional[int]:
        """
        Get the clipboard sequence number.
        
        Returns:
            Sequence number or None if failed
        """
        try:
            import win32clipboard
            
            seq_num = win32clipboard.GetClipboardSequenceNumber()
            logger.info(f"Clipboard sequence number: {seq_num}")
            return seq_num
        except ImportError:
            logger.warning("win32clipboard not available, cannot get clipboard sequence number")
            return None
        except Exception as e:
            logger.error(f"Error getting clipboard sequence number: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def monitor_clipboard_changes(self, callback, timeout=None) -> bool:
        """
        Monitor clipboard changes.
        
        Args:
            callback: Function to call when clipboard changes
            timeout: Timeout in seconds (None for indefinite)
            
        Returns:
            True if successful, False otherwise
        """
        try:
            import win32clipboard
            
            logger.info(f"Monitoring clipboard changes for {timeout if timeout else 'indefinite'} seconds")
            
            last_seq_num = win32clipboard.GetClipboardSequenceNumber()
            start_time = time.time()
            
            while timeout is None or time.time() - start_time < timeout:
                current_seq_num = win32clipboard.GetClipboardSequenceNumber()
                
                if current_seq_num != last_seq_num:
                    logger.info(f"Clipboard changed (sequence number: {current_seq_num})")
                    callback()
                    last_seq_num = current_seq_num
                
                time.sleep(0.5)
            
            logger.info("Clipboard monitoring finished")
            return True
        except ImportError:
            logger.warning("win32clipboard not available, cannot monitor clipboard changes")
            return False
        except Exception as e:
            logger.error(f"Error monitoring clipboard changes: {e}")
            logger.error(traceback.format_exc())
            return False
    def get_screen_resolution(self) -> Tuple[int, int]:
        """
        Get the screen resolution.
        
        Returns:
            Tuple of (width, height)
        """
        try:
            import ctypes
            user32 = ctypes.windll.user32
            width = user32.GetSystemMetrics(0)
            height = user32.GetSystemMetrics(1)
            logger.info(f"Screen resolution: {width}x{height}")
            return (width, height)
        except Exception as e:
            logger.error(f"Error getting screen resolution: {e}")
            logger.error(traceback.format_exc())
            return (1920, 1080)  # Default fallback
    
    def get_screen_dpi(self) -> int:
        """
        Get the screen DPI.
        
        Returns:
            DPI value
        """
        try:
            import ctypes
            user32 = ctypes.windll.user32
            try:
                # Try to get DPI using GetDpiForSystem (Windows 10+)
                dpi = user32.GetDpiForSystem()
            except AttributeError:
                # Fallback for older Windows versions
                dc = user32.GetDC(0)
                dpi = ctypes.windll.gdi32.GetDeviceCaps(dc, 88)  # LOGPIXELSX
                user32.ReleaseDC(0, dc)
            
            logger.info(f"Screen DPI: {dpi}")
            return dpi
        except Exception as e:
            logger.error(f"Error getting screen DPI: {e}")
            logger.error(traceback.format_exc())
            return 96  # Default fallback
    
    def get_screen_scaling_factor(self) -> float:
        """
        Get the screen scaling factor.
        
        Returns:
            Scaling factor (1.0 = 100%, 1.25 = 125%, etc.)
        """
        try:
            dpi = self.get_screen_dpi()
            scaling_factor = dpi / 96.0  # 96 DPI is the base (100%)
            logger.info(f"Screen scaling factor: {scaling_factor}")
            return scaling_factor
        except Exception as e:
            logger.error(f"Error getting screen scaling factor: {e}")
            logger.error(traceback.format_exc())
            return 1.0  # Default fallback
    
    def get_screen_work_area(self) -> Tuple[int, int, int, int]:
        """
        Get the screen work area (excluding taskbar).
        
        Returns:
            Tuple of (left, top, right, bottom)
        """
        try:
            import ctypes
            from ctypes.wintypes import RECT
            
            rect = RECT()
            ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0)  # SPI_GETWORKAREA = 48
            
            work_area = (rect.left, rect.top, rect.right, rect.bottom)
            logger.info(f"Screen work area: {work_area}")
            return work_area
        except Exception as e:
            logger.error(f"Error getting screen work area: {e}")
            logger.error(traceback.format_exc())
            
            # Fallback: Use screen resolution with estimated taskbar size
            width, height = self.get_screen_resolution()
            return (0, 0, width, height - 40)  # Assume taskbar is 40 pixels high
    
    def get_screen_count(self) -> int:
        """
        Get the number of screens.
        
        Returns:
            Number of screens
        """
        try:
            import ctypes
            count = ctypes.windll.user32.GetSystemMetrics(80)  # SM_CMONITORS = 80
            logger.info(f"Screen count: {count}")
            return count
        except Exception as e:
            logger.error(f"Error getting screen count: {e}")
            logger.error(traceback.format_exc())
            return 1  # Default fallback
    
    def get_screen_info(self) -> List[Dict[str, Any]]:
        """
        Get information about all screens.
        
        Returns:
            List of screen information dictionaries
        """
        try:
            import ctypes
            from ctypes.wintypes import RECT
            
            class MONITORINFO(ctypes.Structure):
                _fields_ = [
                    ('cbSize', ctypes.c_ulong),
                    ('rcMonitor', RECT),
                    ('rcWork', RECT),
                    ('dwFlags', ctypes.c_ulong)
                ]
            
            def callback(hMonitor, hdcMonitor, lprcMonitor, dwData):
                info = MONITORINFO()
                info.cbSize = ctypes.sizeof(MONITORINFO)
                ctypes.windll.user32.GetMonitorInfoW(hMonitor, ctypes.byref(info))
                
                screen_info = {
                    'monitor': {
                        'left': info.rcMonitor.left,
                        'top': info.rcMonitor.top,
                        'right': info.rcMonitor.right,
                        'bottom': info.rcMonitor.bottom,
                        'width': info.rcMonitor.right - info.rcMonitor.left,
                        'height': info.rcMonitor.bottom - info.rcMonitor.top
                    },
                    'work_area': {
                        'left': info.rcWork.left,
                        'top': info.rcWork.top,
                        'right': info.rcWork.right,
                        'bottom': info.rcWork.bottom,
                        'width': info.rcWork.right - info.rcWork.left,
                        'height': info.rcWork.bottom - info.rcWork.top
                    },
                    'is_primary': (info.dwFlags & 1) == 1  # MONITORINFOF_PRIMARY = 1
                }
                
                screens.append(screen_info)
                return True
            
            screens = []
            callback_type = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p, ctypes.POINTER(RECT), ctypes.c_void_p)
            callback_func = callback_type(callback)
            
            ctypes.windll.user32.EnumDisplayMonitors(None, None, callback_func, 0)
            
            logger.info(f"Got information for {len(screens)} screens")
            return screens
        except Exception as e:
            logger.error(f"Error getting screen information: {e}")
            logger.error(traceback.format_exc())
            
            # Fallback: Return information for a single screen
            width, height = self.get_screen_resolution()
            work_area = self.get_screen_work_area()
            
            return [{
                'monitor': {
                    'left': 0,
                    'top': 0,
                    'right': width,
                    'bottom': height,
                    'width': width,
                    'height': height
                },
                'work_area': {
                    'left': work_area[0],
                    'top': work_area[1],
                    'right': work_area[2],
                    'bottom': work_area[3],
                    'width': work_area[2] - work_area[0],
                    'height': work_area[3] - work_area[1]
                },
                'is_primary': True
            }]
    
    def get_primary_screen_info(self) -> Dict[str, Any]:
        """
        Get information about the primary screen.
        
        Returns:
            Dictionary with primary screen information
        """
        try:
            screens = self.get_screen_info()
            for screen in screens:
                if screen['is_primary']:
                    logger.info("Got primary screen information")
                    return screen
            
            # If no primary screen found, return the first one
            if screens:
                logger.warning("No primary screen found, returning first screen")
                return screens[0]
            else:
                logger.warning("No screens found")
                return {}
        except Exception as e:
            logger.error(f"Error getting primary screen information: {e}")
            logger.error(traceback.format_exc())
            return {}
    
    def get_cursor_position(self) -> Tuple[int, int]:
        """
        Get the current cursor position.
        
        Returns:
            Tuple of (x, y) coordinates
        """
        try:
            import ctypes
            from ctypes.wintypes import POINT
            
            pt = POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
            
            logger.debug(f"Cursor position: ({pt.x}, {pt.y})")
            return (pt.x, pt.y)
        except Exception as e:
            logger.error(f"Error getting cursor position: {e}")
            logger.error(traceback.format_exc())
            return (0, 0)  # Default fallback
    
    def set_cursor_position(self, x: int, y: int) -> bool:
        """
        Set the cursor position.
        
        Args:
            x: X coordinate
            y: Y coordinate
            
        Returns:
            True if successful, False otherwise
        """
        try:
            import ctypes
            
            result = ctypes.windll.user32.SetCursorPos(x, y)
            
            if result:
                logger.info(f"Set cursor position to ({x}, {y})")
                return True
            else:
                logger.warning(f"Failed to set cursor position to ({x}, {y})")
                return False
        except Exception as e:
            logger.error(f"Error setting cursor position: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_cursor_info(self) -> Dict[str, Any]:
        """
        Get information about the cursor.
        
        Returns:
            Dictionary with cursor information
        """
        try:
            import ctypes
            from ctypes.wintypes import POINT
            
            class CURSORINFO(ctypes.Structure):
                _fields_ = [
                    ('cbSize', ctypes.c_uint),
                    ('flags', ctypes.c_uint),
                    ('hCursor', ctypes.c_void_p),
                    ('ptScreenPos', POINT)
                ]
            
            ci = CURSORINFO()
            ci.cbSize = ctypes.sizeof(CURSORINFO)
            
            if ctypes.windll.user32.GetCursorInfo(ctypes.byref(ci)):
                cursor_info = {
                    'flags': ci.flags,
                    'handle': ci.hCursor,
                    'position': (ci.ptScreenPos.x, ci.ptScreenPos.y),
                    'visible': (ci.flags & 1) == 1  # CURSOR_SHOWING = 1
                }
                
                logger.info(f"Got cursor information: {cursor_info}")
                return cursor_info
            else:
                logger.warning("Failed to get cursor information")
                return {}
        except Exception as e:
            logger.error(f"Error getting cursor information: {e}")
            logger.error(traceback.format_exc())
            return {}
    def get_foreground_window(self) -> Optional[Any]:
        """
        Get the foreground window.
        
        Returns:
            Foreground window object or None if failed
        """
        try:
            import ctypes
            
            hwnd = ctypes.windll.user32.GetForegroundWindow()
            if hwnd:
                # Convert hwnd to pywinauto window
                from pywinauto import Desktop
                desktop = Desktop(backend="uia")
                
                for window in desktop.windows():
                    try:
                        if window.handle == hwnd:
                            logger.info(f"Found foreground window: '{window.window_text()}'")
                            return window
                    except Exception:
                        pass
                
                logger.warning("Could not find pywinauto window for foreground hwnd")
                return None
            else:
                logger.warning("No foreground window")
                return None
        except Exception as e:
            logger.error(f"Error getting foreground window: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def set_foreground_window(self, window) -> bool:
        """
        Set a window as the foreground window.
        
        Args:
            window: Window object
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if hasattr(window, 'set_focus'):
                window.set_focus()
                logger.info(f"Set focus to window: '{window.window_text()}'")
                return True
            else:
                logger.warning("Window doesn't have set_focus method")
                return False
        except Exception as e:
            logger.error(f"Error setting foreground window: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_desktop_windows(self) -> List[Any]:
        """
        Get all desktop windows.
        
        Returns:
            List of window objects
        """
        try:
            from pywinauto import Desktop
            desktop = Desktop(backend="uia")
            
            windows = desktop.windows()
            logger.info(f"Found {len(windows)} desktop windows")
            return windows
        except Exception as e:
            logger.error(f"Error getting desktop windows: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_visible_windows(self) -> List[Any]:
        """
        Get all visible windows.
        
        Returns:
            List of visible window objects
        """
        try:
            from pywinauto import Desktop
            desktop = Desktop(backend="uia")
            
            visible_windows = []
            for window in desktop.windows():
                try:
                    if window.is_visible():
                        visible_windows.append(window)
                except Exception:
                    pass
            
            logger.info(f"Found {len(visible_windows)} visible windows")
            return visible_windows
        except Exception as e:
            logger.error(f"Error getting visible windows: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def get_window_at_position(self, x: int, y: int) -> Optional[Any]:
        """
        Get the window at a specific position.
        
        Args:
            x: X coordinate
            y: Y coordinate
            
        Returns:
            Window object or None if not found
        """
        try:
            import ctypes
            from ctypes.wintypes import POINT
            
            pt = POINT(x, y)
            hwnd = ctypes.windll.user32.WindowFromPoint(pt)
            
            if hwnd:
                # Convert hwnd to pywinauto window
                from pywinauto import Desktop
                desktop = Desktop(backend="uia")
                
                for window in desktop.windows():
                    try:
                        if window.handle == hwnd:
                            logger.info(f"Found window at position ({x}, {y}): '{window.window_text()}'")
                            return window
                    except Exception:
                        pass
                
                logger.warning(f"Could not find pywinauto window for hwnd at position ({x}, {y})")
                return None
            else:
                logger.warning(f"No window at position ({x}, {y})")
                return None
        except Exception as e:
            logger.error(f"Error getting window at position: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_element_at_position(self, x: int, y: int) -> Optional[Any]:
        """
        Get the element at a specific position.
        
        Args:
            x: X coordinate
            y: Y coordinate
            
        Returns:
            Element object or None if not found
        """
        try:
            # First get the window at the position
            window = self.get_window_at_position(x, y)
            if not window:
                return None
            
            # Then get the element within the window
            if hasattr(window, 'from_point'):
                element = window.from_point(x, y)
                if element:
                    logger.info(f"Found element at position ({x}, {y}): '{element.window_text() if hasattr(element, 'window_text') else 'Unknown'}'")
                    return element
                else:
                    logger.warning(f"No element at position ({x}, {y}) within window")
                    return None
            else:
                logger.warning("Window doesn't have from_point method")
                return None
        except Exception as e:
            logger.error(f"Error getting element at position: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def capture_screen(self, file_path: str = None, region: Tuple[int, int, int, int] = None) -> Optional[Any]:
        """
        Capture the screen or a region of the screen.
        
        Args:
            file_path: Path to save the image (if None, image is returned but not saved)
            region: Region to capture as (left, top, width, height) (if None, entire screen is captured)
            
        Returns:
            Image object or None if failed
        """
        try:
            from PIL import ImageGrab
            
            if region:
                left, top, width, height = region
                right = left + width
                bottom = top + height
                image = ImageGrab.grab(bbox=(left, top, right, bottom))
            else:
                image = ImageGrab.grab()
            
            if file_path:
                image.save(file_path)
                logger.info(f"Captured screen to file: {file_path}")
            else:
                logger.info("Captured screen to memory")
            
            return image
        except ImportError:
            logger.warning("PIL not available, cannot capture screen")
            return None
        except Exception as e:
            logger.error(f"Error capturing screen: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def capture_window(self, window, file_path: str = None) -> Optional[Any]:
        """
        Capture a window.
        
        Args:
            window: Window object
            file_path: Path to save the image (if None, image is returned but not saved)
            
        Returns:
            Image object or None if failed
        """
        try:
            if hasattr(window, 'capture_as_image'):
                image = window.capture_as_image()
                
                if file_path:
                    image.save(file_path)
                    logger.info(f"Captured window to file: {file_path}")
                else:
                    logger.info("Captured window to memory")
                
                return image
            else:
                logger.warning("Window doesn't have capture_as_image method")
                return None
        except Exception as e:
            logger.error(f"Error capturing window: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def capture_element(self, element, file_path: str = None) -> Optional[Any]:
        """
        Capture an element.
        
        Args:
            element: Element object
            file_path: Path to save the image (if None, image is returned but not saved)
            
        Returns:
            Image object or None if failed
        """
        try:
            if hasattr(element, 'capture_as_image'):
                image = element.capture_as_image()
                
                if file_path:
                    image.save(file_path)
                    logger.info(f"Captured element to file: {file_path}")
                else:
                    logger.info("Captured element to memory")
                
                return image
            else:
                logger.warning("Element doesn't have capture_as_image method")
                return None
        except Exception as e:
            logger.error(f"Error capturing element: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def get_pixel_color(self, x: int, y: int) -> Optional[Tuple[int, int, int]]:
        """
        Get the color of a pixel at a specific position.
        
        Args:
            x: X coordinate
            y: Y coordinate
            
        Returns:
            Tuple of (R, G, B) values or None if failed
        """
        try:
            from PIL import ImageGrab
            
            image = ImageGrab.grab(bbox=(x, y, x+1, y+1))
            color = image.getpixel((0, 0))
            
            logger.debug(f"Pixel color at ({x}, {y}): {color}")
            return color
        except ImportError:
            logger.warning("PIL not available, cannot get pixel color")
            return None
        except Exception as e:
            logger.error(f"Error getting pixel color: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def find_image_on_screen(self, image_path: str, confidence: float = 0.9, region: Tuple[int, int, int, int] = None) -> Optional[Tuple[int, int]]:
        """
        Find an image on the screen.
        
        Args:
            image_path: Path to the image file
            confidence: Confidence threshold (0.0 to 1.0)
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            
        Returns:
            Tuple of (x, y) coordinates of the center of the found image or None if not found
        """
        try:
            import cv2
            import numpy as np
            from PIL import ImageGrab
            
            # Load the template image
            template = cv2.imread(image_path)
            if template is None:
                logger.error(f"Failed to load image: {image_path}")
                return None
            
            # Capture the screen
            if region:
                left, top, width, height = region
                right = left + width
                bottom = top + height
                screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
            else:
                screenshot = ImageGrab.grab()
                left, top = 0, 0
            
            # Convert PIL image to OpenCV format
            screenshot = np.array(screenshot)
            screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)
            
            # Perform template matching
            result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            if max_val >= confidence:
                # Get the center of the found image
                h, w = template.shape[:2]
                center_x = max_loc[0] + w // 2 + left
                center_y = max_loc[1] + h // 2 + top
                
                logger.info(f"Found image at position ({center_x}, {center_y}) with confidence {max_val:.2f}")
                return (center_x, center_y)
            else:
                logger.warning(f"Image not found (max confidence: {max_val:.2f})")
                return None
        except ImportError:
            logger.warning("OpenCV or PIL not available, cannot find image on screen")
            return None
        except Exception as e:
            logger.error(f"Error finding image on screen: {e}")
            logger.error(traceback.format_exc())
            return None
    def find_all_images_on_screen(self, image_path: str, confidence: float = 0.9, region: Tuple[int, int, int, int] = None) -> List[Tuple[int, int]]:
        """
        Find all occurrences of an image on the screen.
        
        Args:
            image_path: Path to the image file
            confidence: Confidence threshold (0.0 to 1.0)
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            
        Returns:
            List of (x, y) coordinates of the centers of the found images
        """
        try:
            import cv2
            import numpy as np
            from PIL import ImageGrab
            
            # Load the template image
            template = cv2.imread(image_path)
            if template is None:
                logger.error(f"Failed to load image: {image_path}")
                return []
            
            # Capture the screen
            if region:
                left, top, width, height = region
                right = left + width
                bottom = top + height
                screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
            else:
                screenshot = ImageGrab.grab()
                left, top = 0, 0
            
            # Convert PIL image to OpenCV format
            screenshot = np.array(screenshot)
            screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)
            
            # Perform template matching
            result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
            
            # Find all matches above the confidence threshold
            locations = np.where(result >= confidence)
            
            # Get the centers of the found images
            h, w = template.shape[:2]
            centers = []
            
            for pt in zip(*locations[::-1]):
                center_x = pt[0] + w // 2 + left
                center_y = pt[1] + h // 2 + top
                centers.append((center_x, center_y))
            
            logger.info(f"Found {len(centers)} images with confidence >= {confidence:.2f}")
            return centers
        except ImportError:
            logger.warning("OpenCV or PIL not available, cannot find images on screen")
            return []
        except Exception as e:
            logger.error(f"Error finding images on screen: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def wait_for_image(self, image_path: str, confidence: float = 0.9, region: Tuple[int, int, int, int] = None, timeout: int = None) -> Optional[Tuple[int, int]]:
        """
        Wait for an image to appear on the screen.
        
        Args:
            image_path: Path to the image file
            confidence: Confidence threshold (0.0 to 1.0)
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            timeout: Timeout in seconds
            
        Returns:
            Tuple of (x, y) coordinates of the center of the found image or None if not found
        """
        timeout = timeout or self.default_timeout
        
        logger.info(f"Waiting for image: {image_path} (timeout: {timeout}s)")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            result = self.find_image_on_screen(image_path, confidence, region)
            if result:
                return result
            time.sleep(0.5)
        
        logger.warning(f"Image not found after {timeout} seconds")
        return None
    
    def click_image(self, image_path: str, confidence: float = 0.9, region: Tuple[int, int, int, int] = None, timeout: int = None) -> bool:
        """
        Click on an image on the screen.
        
        Args:
            image_path: Path to the image file
            confidence: Confidence threshold (0.0 to 1.0)
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            timeout: Timeout in seconds
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Wait for the image to appear
            position = self.wait_for_image(image_path, confidence, region, timeout)
            if not position:
                return False
            
            # Click on the image
            x, y = position
            self.click_at_coordinates(x, y)
            
            logger.info(f"Clicked on image at position ({x}, {y})")
            return True
        except Exception as e:
            logger.error(f"Error clicking on image: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def double_click_image(self, image_path: str, confidence: float = 0.9, region: Tuple[int, int, int, int] = None, timeout: int = None) -> bool:
        """
        Double-click on an image on the screen.
        
        Args:
            image_path: Path to the image file
            confidence: Confidence threshold (0.0 to 1.0)
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            timeout: Timeout in seconds
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Wait for the image to appear
            position = self.wait_for_image(image_path, confidence, region, timeout)
            if not position:
                return False
            
            # Double-click on the image
            x, y = position
            self.double_click_at_coordinates(x, y)
            
            logger.info(f"Double-clicked on image at position ({x}, {y})")
            return True
        except Exception as e:
            logger.error(f"Error double-clicking on image: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def right_click_image(self, image_path: str, confidence: float = 0.9, region: Tuple[int, int, int, int] = None, timeout: int = None) -> bool:
        """
        Right-click on an image on the screen.
        
        Args:
            image_path: Path to the image file
            confidence: Confidence threshold (0.0 to 1.0)
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            timeout: Timeout in seconds
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Wait for the image to appear
            position = self.wait_for_image(image_path, confidence, region, timeout)
            if not position:
                return False
            
            # Right-click on the image
            x, y = position
            self.right_click_at_coordinates(x, y)
            
            logger.info(f"Right-clicked on image at position ({x}, {y})")
            return True
        except Exception as e:
            logger.error(f"Error right-clicking on image: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def wait_for_image_to_disappear(self, image_path: str, confidence: float = 0.9, region: Tuple[int, int, int, int] = None, timeout: int = None) -> bool:
        """
        Wait for an image to disappear from the screen.
        
        Args:
            image_path: Path to the image file
            confidence: Confidence threshold (0.0 to 1.0)
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            timeout: Timeout in seconds
            
        Returns:
            True if image disappears, False otherwise
        """
        timeout = timeout or self.default_timeout
        
        logger.info(f"Waiting for image to disappear: {image_path} (timeout: {timeout}s)")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            result = self.find_image_on_screen(image_path, confidence, region)
            if not result:
                logger.info("Image disappeared")
                return True
            time.sleep(0.5)
        
        logger.warning(f"Image still present after {timeout} seconds")
        return False
    
    def wait_for_any_image(self, image_paths: List[str], confidence: float = 0.9, region: Tuple[int, int, int, int] = None, timeout: int = None) -> Tuple[Optional[str], Optional[Tuple[int, int]]]:
        """
        Wait for any of several images to appear on the screen.
        
        Args:
            image_paths: List of paths to image files
            confidence: Confidence threshold (0.0 to 1.0)
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            timeout: Timeout in seconds
            
        Returns:
            Tuple of (image_path, position) for the first image found, or (None, None) if none found
        """
        timeout = timeout or self.default_timeout
        
        logger.info(f"Waiting for any of {len(image_paths)} images (timeout: {timeout}s)")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            for image_path in image_paths:
                result = self.find_image_on_screen(image_path, confidence, region)
                if result:
                    logger.info(f"Found image: {image_path} at position {result}")
                    return (image_path, result)
            time.sleep(0.5)
        
        logger.warning(f"No images found after {timeout} seconds")
        return (None, None)
    
    def click_any_image(self, image_paths: List[str], confidence: float = 0.9, region: Tuple[int, int, int, int] = None, timeout: int = None) -> Optional[str]:
        """
        Click on any of several images on the screen.
        
        Args:
            image_paths: List of paths to image files
            confidence: Confidence threshold (0.0 to 1.0)
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            timeout: Timeout in seconds
            
        Returns:
            Path of the clicked image or None if none found
        """
        try:
            # Wait for any image to appear
            image_path, position = self.wait_for_any_image(image_paths, confidence, region, timeout)
            if not image_path or not position:
                return None
            
            # Click on the image
            x, y = position
            self.click_at_coordinates(x, y)
            
            logger.info(f"Clicked on image {image_path} at position ({x}, {y})")
            return image_path
        except Exception as e:
            logger.error(f"Error clicking on any image: {e}")
            logger.error(traceback.format_exc())
            return None
    def compare_images(self, image_path1: str, image_path2: str) -> float:
        """
        Compare two images and return their similarity.
        
        Args:
            image_path1: Path to the first image file
            image_path2: Path to the second image file
            
        Returns:
            Similarity score (0.0 to 1.0)
        """
        try:
            import cv2
            import numpy as np
            
            # Load the images
            img1 = cv2.imread(image_path1)
            img2 = cv2.imread(image_path2)
            
            if img1 is None or img2 is None:
                logger.error("Failed to load one or both images")
                return 0.0
            
            # Resize images to the same size
            height = min(img1.shape[0], img2.shape[0])
            width = min(img1.shape[1], img2.shape[1])
            img1 = cv2.resize(img1, (width, height))
            img2 = cv2.resize(img2, (width, height))
            
            # Convert images to grayscale
            gray1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
            gray2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
            
            # Calculate structural similarity index
            (score, diff) = cv2.compareSSIM(gray1, gray2, full=True)
            
            logger.info(f"Image similarity score: {score:.4f}")
            return score
        except ImportError:
            logger.warning("OpenCV not available, cannot compare images")
            return 0.0
        except Exception as e:
            logger.error(f"Error comparing images: {e}")
            logger.error(traceback.format_exc())
            return 0.0
    
    def is_image_on_screen(self, image_path: str, confidence: float = 0.9, region: Tuple[int, int, int, int] = None) -> bool:
        """
        Check if an image is on the screen.
        
        Args:
            image_path: Path to the image file
            confidence: Confidence threshold (0.0 to 1.0)
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            
        Returns:
            True if image is found, False otherwise
        """
        result = self.find_image_on_screen(image_path, confidence, region)
        return result is not None
    
    def count_images_on_screen(self, image_path: str, confidence: float = 0.9, region: Tuple[int, int, int, int] = None) -> int:
        """
        Count the number of occurrences of an image on the screen.
        
        Args:
            image_path: Path to the image file
            confidence: Confidence threshold (0.0 to 1.0)
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            
        Returns:
            Number of occurrences
        """
        results = self.find_all_images_on_screen(image_path, confidence, region)
        return len(results)
    
    def get_text_from_image(self, image_path: str = None, region: Tuple[int, int, int, int] = None) -> str:
        """
        Extract text from an image or screen region using OCR.
        
        Args:
            image_path: Path to the image file (if None, capture screen region)
            region: Region to capture as (left, top, width, height) (if None and image_path is None, entire screen is captured)
            
        Returns:
            Extracted text
        """
        try:
            import pytesseract
            from PIL import Image
            
            # Get the image
            if image_path:
                image = Image.open(image_path)
            else:
                image = self.capture_screen(file_path=None, region=region)
                if image is None:
                    logger.error("Failed to capture screen")
                    return ""
            
            # Extract text using Tesseract OCR
            text = pytesseract.image_to_string(image)
            
            logger.info(f"Extracted text from image: {text[:50]}...")
            return text
        except ImportError:
            logger.warning("pytesseract not available, cannot extract text from image")
            return ""
        except Exception as e:
            logger.error(f"Error extracting text from image: {e}")
            logger.error(traceback.format_exc())
            return ""
    
    def find_text_on_screen(self, text: str, region: Tuple[int, int, int, int] = None) -> Optional[Tuple[int, int, int, int]]:
        """
        Find text on the screen using OCR.
        
        Args:
            text: Text to find
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            
        Returns:
            Tuple of (left, top, right, bottom) coordinates of the found text or None if not found
        """
        try:
            import pytesseract
            from PIL import Image
            
            # Capture the screen
            image = self.capture_screen(file_path=None, region=region)
            if image is None:
                logger.error("Failed to capture screen")
                return None
            
            # Extract text and bounding boxes using Tesseract OCR
            data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
            
            # Calculate region offset
            left_offset = 0
            top_offset = 0
            if region:
                left_offset = region[0]
                top_offset = region[1]
            
            # Search for the text
            for i in range(len(data['text'])):
                if data['text'][i].strip() == text:
                    x = data['left'][i] + left_offset
                    y = data['top'][i] + top_offset
                    w = data['width'][i]
                    h = data['height'][i]
                    
                    logger.info(f"Found text '{text}' at position ({x}, {y}, {x+w}, {y+h})")
                    return (x, y, x+w, y+h)
            
            logger.warning(f"Text '{text}' not found on screen")
            return None
        except ImportError:
            logger.warning("pytesseract not available, cannot find text on screen")
            return None
        except Exception as e:
            logger.error(f"Error finding text on screen: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def click_text(self, text: str, region: Tuple[int, int, int, int] = None) -> bool:
        """
        Click on text on the screen.
        
        Args:
            text: Text to click
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Find the text on the screen
            bbox = self.find_text_on_screen(text, region)
            if not bbox:
                return False
            
            # Calculate the center of the bounding box
            x = (bbox[0] + bbox[2]) // 2
            y = (bbox[1] + bbox[3]) // 2
            
            # Click on the text
            self.click_at_coordinates(x, y)
            
            logger.info(f"Clicked on text '{text}' at position ({x}, {y})")
            return True
        except Exception as e:
            logger.error(f"Error clicking on text: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def wait_for_text(self, text: str, region: Tuple[int, int, int, int] = None, timeout: int = None) -> Optional[Tuple[int, int, int, int]]:
        """
        Wait for text to appear on the screen.
        
        Args:
            text: Text to wait for
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            timeout: Timeout in seconds
            
        Returns:
            Tuple of (left, top, right, bottom) coordinates of the found text or None if not found
        """
        timeout = timeout or self.default_timeout
        
        logger.info(f"Waiting for text: '{text}' (timeout: {timeout}s)")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            result = self.find_text_on_screen(text, region)
            if result:
                return result
            time.sleep(0.5)
        
        logger.warning(f"Text '{text}' not found after {timeout} seconds")
        return None
    
    def wait_for_text_to_disappear(self, text: str, region: Tuple[int, int, int, int] = None, timeout: int = None) -> bool:
        """
        Wait for text to disappear from the screen.
        
        Args:
            text: Text to wait for
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            timeout: Timeout in seconds
            
        Returns:
            True if text disappears, False otherwise
        """
        timeout = timeout or self.default_timeout
        
        logger.info(f"Waiting for text to disappear: '{text}' (timeout: {timeout}s)")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            result = self.find_text_on_screen(text, region)
            if not result:
                logger.info("Text disappeared")
                return True
            time.sleep(0.5)
        
        logger.warning(f"Text '{text}' still present after {timeout} seconds")
        return False
    
    def is_text_on_screen(self, text: str, region: Tuple[int, int, int, int] = None) -> bool:
        """
        Check if text is on the screen.
        
        Args:
            text: Text to check
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            
        Returns:
            True if text is found, False otherwise
        """
        result = self.find_text_on_screen(text, region)
        return result is not None
    def get_all_text_on_screen(self, region: Tuple[int, int, int, int] = None) -> str:
        """
        Get all text on the screen using OCR.
        
        Args:
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            
        Returns:
            All text found on the screen
        """
        return self.get_text_from_image(None, region)
    
    def find_all_text_on_screen(self, region: Tuple[int, int, int, int] = None) -> List[Dict[str, Any]]:
        """
        Find all text on the screen using OCR.
        
        Args:
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            
        Returns:
            List of dictionaries with text and bounding box information
        """
        try:
            import pytesseract
            from PIL import Image
            
            # Capture the screen
            image = self.capture_screen(file_path=None, region=region)
            if image is None:
                logger.error("Failed to capture screen")
                return []
            
            # Extract text and bounding boxes using Tesseract OCR
            data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
            
            # Calculate region offset
            left_offset = 0
            top_offset = 0
            if region:
                left_offset = region[0]
                top_offset = region[1]
            
            # Collect all text with bounding boxes
            results = []
            for i in range(len(data['text'])):
                if data['text'][i].strip():
                    x = data['left'][i] + left_offset
                    y = data['top'][i] + top_offset
                    w = data['width'][i]
                    h = data['height'][i]
                    
                    results.append({
                        'text': data['text'][i],
                        'bbox': (x, y, x+w, y+h),
                        'conf': data['conf'][i]
                    })
            
            logger.info(f"Found {len(results)} text elements on screen")
            return results
        except ImportError:
            logger.warning("pytesseract not available, cannot find all text on screen")
            return []
        except Exception as e:
            logger.error(f"Error finding all text on screen: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def find_text_matching_pattern(self, pattern: str, region: Tuple[int, int, int, int] = None) -> List[Dict[str, Any]]:
        """
        Find all text on the screen matching a pattern.
        
        Args:
            pattern: Regular expression pattern
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            
        Returns:
            List of dictionaries with text and bounding box information
        """
        try:
            import re
            
            # Get all text on the screen
            all_text = self.find_all_text_on_screen(region)
            
            # Filter by pattern
            pattern_obj = re.compile(pattern)
            matching_text = [item for item in all_text if pattern_obj.search(item['text'])]
            
            logger.info(f"Found {len(matching_text)} text elements matching pattern '{pattern}'")
            return matching_text
        except Exception as e:
            logger.error(f"Error finding text matching pattern: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def click_text_matching_pattern(self, pattern: str, region: Tuple[int, int, int, int] = None) -> bool:
        """
        Click on text on the screen matching a pattern.
        
        Args:
            pattern: Regular expression pattern
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Find text matching the pattern
            matching_text = self.find_text_matching_pattern(pattern, region)
            if not matching_text:
                return False
            
            # Click on the first match
            bbox = matching_text[0]['bbox']
            x = (bbox[0] + bbox[2]) // 2
            y = (bbox[1] + bbox[3]) // 2
            
            self.click_at_coordinates(x, y)
            
            logger.info(f"Clicked on text '{matching_text[0]['text']}' matching pattern '{pattern}' at position ({x}, {y})")
            return True
        except Exception as e:
            logger.error(f"Error clicking on text matching pattern: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def wait_for_text_matching_pattern(self, pattern: str, region: Tuple[int, int, int, int] = None, timeout: int = None) -> Optional[Dict[str, Any]]:
        """
        Wait for text matching a pattern to appear on the screen.
        
        Args:
            pattern: Regular expression pattern
            region: Region to search in as (left, top, width, height) (if None, entire screen is searched)
            timeout: Timeout in seconds
            
        Returns:
            Dictionary with text and bounding box information or None if not found
        """
        timeout = timeout or self.default_timeout
        
        logger.info(f"Waiting for text matching pattern: '{pattern}' (timeout: {timeout}s)")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            matching_text = self.find_text_matching_pattern(pattern, region)
            if matching_text:
                logger.info(f"Found text '{matching_text[0]['text']}' matching pattern '{pattern}'")
                return matching_text[0]
            time.sleep(0.5)
        
        logger.warning(f"Text matching pattern '{pattern}' not found after {timeout} seconds")
        return None
    
    def get_text_from_element(self, element) -> str:
        """
        Get text from an element.
        
        Args:
            element: Element object
            
        Returns:
            Element text
        """
        try:
            if hasattr(element, 'window_text'):
                text = element.window_text()
                logger.debug(f"Got text from element: {text}")
                return text
            else:
                logger.warning("Element doesn't have window_text method")
                return ""
        except Exception as e:
            logger.error(f"Error getting text from element: {e}")
            logger.error(traceback.format_exc())
            return ""
    
    def get_all_text_from_elements(self, elements: List[Any]) -> List[str]:
        """
        Get text from multiple elements.
        
        Args:
            elements: List of element objects
            
        Returns:
            List of element texts
        """
        texts = []
        for element in elements:
            text = self.get_text_from_element(element)
            if text:
                texts.append(text)
        
        logger.debug(f"Got text from {len(texts)} elements")
        return texts
    
    def find_element_by_text_content(self, parent_window, text: str, partial_match: bool = False) -> Optional[Any]:
        """
        Find an element by its text content.
        
        Args:
            parent_window: Parent window object
            text: Text to search for
            partial_match: Whether to allow partial matches
            
        Returns:
            Element object or None if not found
        """
        try:
            # Get all descendants
            descendants = parent_window.descendants()
            
            for element in descendants:
                try:
                    if hasattr(element, 'window_text'):
                        element_text = element.window_text()
                        
                        if (partial_match and text in element_text) or (not partial_match and text == element_text):
                            logger.info(f"Found element with text: '{element_text}'")
                            return element
                except Exception:
                    continue
            
            logger.warning(f"Element with text '{text}' not found")
            return None
        except Exception as e:
            logger.error(f"Error finding element by text content: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def find_all_elements_by_text_content(self, parent_window, text: str, partial_match: bool = False) -> List[Any]:
        """
        Find all elements by their text content.
        
        Args:
            parent_window: Parent window object
            text: Text to search for
            partial_match: Whether to allow partial matches
            
        Returns:
            List of element objects
        """
        try:
            # Get all descendants
            descendants = parent_window.descendants()
            
            matching_elements = []
            for element in descendants:
                try:
                    if hasattr(element, 'window_text'):
                        element_text = element.window_text()
                        
                        if (partial_match and text in element_text) or (not partial_match and text == element_text):
                            matching_elements.append(element)
                except Exception:
                    continue
            
            logger.info(f"Found {len(matching_elements)} elements with text '{text}'")
            return matching_elements
        except Exception as e:
            logger.error(f"Error finding elements by text content: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def find_element_by_text_content_regex(self, parent_window, pattern: str) -> Optional[Any]:
        """
        Find an element by its text content using a regular expression.
        
        Args:
            parent_window: Parent window object
            pattern: Regular expression pattern
            
        Returns:
            Element object or None if not found
        """
        try:
            import re
            
            # Compile the pattern
            pattern_obj = re.compile(pattern)
            
            # Get all descendants
            descendants = parent_window.descendants()
            
            for element in descendants:
                try:
                    if hasattr(element, 'window_text'):
                        element_text = element.window_text()
                        
                        if pattern_obj.search(element_text):
                            logger.info(f"Found element with text matching pattern '{pattern}': '{element_text}'")
                            return element
                except Exception:
                    continue
            
            logger.warning(f"Element with text matching pattern '{pattern}' not found")
            return None
        except Exception as e:
            logger.error(f"Error finding element by text content regex: {e}")
            logger.error(traceback.format_exc())
            return None
    def find_all_elements_by_text_content_regex(self, parent_window, pattern: str) -> List[Any]:
        """
        Find all elements by their text content using a regular expression.
        
        Args:
            parent_window: Parent window object
            pattern: Regular expression pattern
            
        Returns:
            List of element objects
        """
        try:
            import re
            
            # Compile the pattern
            pattern_obj = re.compile(pattern)
            
            # Get all descendants
            descendants = parent_window.descendants()
            
            matching_elements = []
            for element in descendants:
                try:
                    if hasattr(element, 'window_text'):
                        element_text = element.window_text()
                        
                        if pattern_obj.search(element_text):
                            matching_elements.append(element)
                except Exception:
                    continue
            
            logger.info(f"Found {len(matching_elements)} elements with text matching pattern '{pattern}'")
            return matching_elements
        except Exception as e:
            logger.error(f"Error finding elements by text content regex: {e}")
            logger.error(traceback.format_exc())
            return []
    
    def click_element_by_text(self, parent_window, text: str, partial_match: bool = False) -> bool:
        """
        Click on an element by its text content.
        
        Args:
            parent_window: Parent window object
            text: Text to search for
            partial_match: Whether to allow partial matches
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Find the element
            element = self.find_element_by_text_content(parent_window, text, partial_match)
            if not element:
                return False
            
            # Click the element
            element.click_input()
            
            logger.info(f"Clicked on element with text: '{text}'")
            return True
        except Exception as e:
            logger.error(f"Error clicking element by text: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def click_element_by_text_regex(self, parent_window, pattern: str) -> bool:
        """
        Click on an element by its text content using a regular expression.
        
        Args:
            parent_window: Parent window object
            pattern: Regular expression pattern
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Find the element
            element = self.find_element_by_text_content_regex(parent_window, pattern)
            if not element:
                return False
            
            # Click the element
            element.click_input()
            
            logger.info(f"Clicked on element with text matching pattern: '{pattern}'")
            return True
        except Exception as e:
            logger.error(f"Error clicking element by text regex: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def wait_for_element_by_text(self, parent_window, text: str, partial_match: bool = False, timeout: int = None) -> Optional[Any]:
        """
        Wait for an element with specific text to appear.
        
        Args:
            parent_window: Parent window object
            text: Text to search for
            partial_match: Whether to allow partial matches
            timeout: Timeout in seconds
            
        Returns:
            Element object or None if not found
        """
        timeout = timeout or self.element_timeout
        
        logger.info(f"Waiting for element with text: '{text}' (timeout: {timeout}s)")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            element = self.find_element_by_text_content(parent_window, text, partial_match)
            if element:
                return element
            time.sleep(0.5)
        
        logger.warning(f"Element with text '{text}' not found after {timeout} seconds")
        return None
    
    def wait_for_element_by_text_regex(self, parent_window, pattern: str, timeout: int = None) -> Optional[Any]:
        """
        Wait for an element with text matching a regular expression to appear.
        
        Args:
            parent_window: Parent window object
            pattern: Regular expression pattern
            timeout: Timeout in seconds
            
        Returns:
            Element object or None if not found
        """
        timeout = timeout or self.element_timeout
        
        logger.info(f"Waiting for element with text matching pattern: '{pattern}' (timeout: {timeout}s)")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            element = self.find_element_by_text_content_regex(parent_window, pattern)
            if element:
                return element
            time.sleep(0.5)
        
        logger.warning(f"Element with text matching pattern '{pattern}' not found after {timeout} seconds")
        return None
    
    def wait_for_element_by_text_to_disappear(self, parent_window, text: str, partial_match: bool = False, timeout: int = None) -> bool:
        """
        Wait for an element with specific text to disappear.
        
        Args:
            parent_window: Parent window object
            text: Text to search for
            partial_match: Whether to allow partial matches
            timeout: Timeout in seconds
            
        Returns:
            True if element disappears, False otherwise
        """
        timeout = timeout or self.element_timeout
        
        logger.info(f"Waiting for element with text '{text}' to disappear (timeout: {timeout}s)")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            element = self.find_element_by_text_content(parent_window, text, partial_match)
            if not element:
                logger.info(f"Element with text '{text}' disappeared")
                return True
            time.sleep(0.5)
        
        logger.warning(f"Element with text '{text}' still present after {timeout} seconds")
        return False
    
    def wait_for_element_by_text_regex_to_disappear(self, parent_window, pattern: str, timeout: int = None) -> bool:
        """
        Wait for an element with text matching a regular expression to disappear.
        
        Args:
            parent_window: Parent window object
            pattern: Regular expression pattern
            timeout: Timeout in seconds
            
        Returns:
            True if element disappears, False otherwise
        """
        timeout = timeout or self.element_timeout
        
        logger.info(f"Waiting for element with text matching pattern '{pattern}' to disappear (timeout: {timeout}s)")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            element = self.find_element_by_text_content_regex(parent_window, pattern)
            if not element:
                logger.info(f"Element with text matching pattern '{pattern}' disappeared")
                return True
            time.sleep(0.5)
        
        logger.warning(f"Element with text matching pattern '{pattern}' still present after {timeout} seconds")
        return False
    
    def is_element_with_text_present(self, parent_window, text: str, partial_match: bool = False) -> bool:
        """
        Check if an element with specific text is present.
        
        Args:
            parent_window: Parent window object
            text: Text to search for
            partial_match: Whether to allow partial matches
            
        Returns:
            True if element is present, False otherwise
        """
        element = self.find_element_by_text_content(parent_window, text, partial_match)
        return element is not None
    
    def is_element_with_text_regex_present(self, parent_window, pattern: str) -> bool:
        """
        Check if an element with text matching a regular expression is present.
        
        Args:
            parent_window: Parent window object
            pattern: Regular expression pattern
            
        Returns:
            True if element is present, False otherwise
        """
        element = self.find_element_by_text_content_regex(parent_window, pattern)
        return element is not None
