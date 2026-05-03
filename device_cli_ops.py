# Standard library module for thread-safe signaling between threads (used for cancellation)
import threading

# Standard library module for interacting with the Python interpreter (used here for sys.exit())
import sys

# Standard library module for low‑level network communication
import socket

# Standard library module to perform operations with time
import time

# Standard library module for managing concurrent execution across multiple threads
from concurrent.futures import ThreadPoolExecutor, as_completed

# Third-party library for SSH connections to network devices and sending CLI commands
from netmiko import ConnectHandler

# Catch auth and timeout failures separately from connectivity failures
from netmiko.exceptions import (
    NetmikoAuthenticationException,
    NetmikoTimeoutException,
    ReadTimeout,
    ConfigInvalidException,
    
)

# Catch timeout failures via Paramiko methods
from paramiko.ssh_exception import SSHException

# Standard library to perform operations with files and folders
import os

# Standard library module for writing structured log messages (INFO, ERROR, etc.)
import logging

# Logging handler that auto-rotates the log file when it reaches a defined size limit
from logging.handlers import RotatingFileHandler

###################################################################################################################################

# SETUP LOGGING
# -------------
logger = logging.getLogger(__name__) # use the module's name as the name in the logs
logger.setLevel(logging.INFO) # set the logging level
log_file_path = 'execution_logs.log' # define the logging file path

# Use RotatingFileHandler
# log_file_path = the absolute path for the log file
# maxBytes: 5 * 1024
# backupCount=0: When the file is full, delete it and start a new one
handler = RotatingFileHandler(
    log_file_path, maxBytes=5*1024*1024, backupCount=0
)

# Set the format of the log messages
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
handler.setFormatter(formatter)

# Add the formatted HANDLER to the logger
logger.addHandler(handler)

###################################################################################################################################

# DEFINE A GLOBAL CANCELLATION EVENT
# ----------------------------------
# Set this to signal all threads to stop
cancel_event = threading.Event()


# CHECK CONNECTIVITY VIA SSH
# --------------------------
def first_connectivity_check(row, username, password):

    """
    PURPOSE
    -------
    Connects to a device via Netmiko and runs a test command.
    Sets the device Status to ONLINE on success, OFFLINE on failure.
    If authentication fails, marks device as ONLINE with AUTH_BAD auth status.

    ARGUMENTS
    ---------
    row      (pd.Series): A single DataFrame row containing device information.
    username (str):       Device SSH username.
    password (str):       Device SSH password.

    RETURN VALUE
    ------------
    Tuple of (index, status, auth_status) where:
        status      : 'ONLINE' or 'OFFLINE'
        auth_status : 'AUTH_OK', 'AUTH_BAD', or None
    """

    # Define the parameters
    ip = row['OOBM IP Address']
    hostname = row['Hostname']
    index = row.name

    # Open the SSH connection
    try:
        connection = ConnectHandler(
            device_type='cisco_ios',
            host=ip,
            username=username,
            password=password
        )

        # Define the test command
        commands = [
            '!'
        ]

        # Push the config commands to the devices
        connection.send_config_set(commands)
        connection.disconnect()

        #--------------------------------------------------------------------
        logger.info(f"Test command sent successfully to '{hostname}' ({ip})")
        #--------------------------------------------------------------------

        # If it succeded, it means the device is ONLINE and the authentication succeded
        return index, 'ONLINE', 'AUTH_OK'
    
    except NetmikoAuthenticationException:
        #------------------------------------------------------------------------------------
        logger.warning(f"Authentication failed on '{hostname}' ({ip}) — device is reachable")
        #------------------------------------------------------------------------------------
        # Device responded but rejected credentials — it's ONLINE but AUTH_BAD
        return index, 'ONLINE', 'AUTH_BAD'
    
    except Exception as e:
        #--------------------------------------------------------------------------
        logger.warning(f"Device '{hostname}' ({ip}) is unreachable or failed: {e}")
        #--------------------------------------------------------------------------
        # If it failed, it means that the device is OFFLINE
        return index, 'OFFLINE', 'N/A'

# Thread pool -- Called by 'populate_status_and_auth_status_column' from 'excel_and_data_ops'
def first_connectivity_check_all(valid_devices_df, username, password):

    """
    PURPOSE
    -------
    Concurrently connects to all devices in the DataFrame and runs a test command.


    ARGUMENTS
    ---------
    valid_devices_df (pd.DataFrame): DataFrame containing at minimum 'OOBM IP Address' and 'Hostname' columns.
    username         (str):          Device SSH username.
    password         (str):          Device SSH password.


    RETURN VALUE
    ------------
    List of (index, status, auth_status) tuples — ONLINE/OFFLINE for status, AUTH_OK/AUTH_BAD/N/A for auth_status.
    """

    # Submit one configuration task per device — max 10 concurrent connections
    with ThreadPoolExecutor(max_workers=10) as executor:

        # Build a future for each device, keyed by hostname for result logging
        futures = {
            executor.submit(first_connectivity_check, row, username, password): row['Hostname']
            for _, row in valid_devices_df.iterrows()
        }

        # Initialize the list that will collect (index, status) tuples
        results = []

        # Process results as each thread completes — order is not guaranteed
        for future in as_completed(futures):
            hostname = futures[future]
            try:
                # Unpack the (index, status, auth_status) tuple returned by 'first_connectivity_check'
                index, status, auth_status = future.result()
                results.append((index, status, auth_status))
                #-------------------------------------------------------------------------
                logger.info(f"'{hostname}' status set to: {status} | auth: {auth_status}")
                #-------------------------------------------------------------------------
            except Exception as e:
                #-------------------------------------------------------------
                logger.error(f"Unexpected error processing '{hostname}': {e}")
                #-------------------------------------------------------------

    return results


# NOT USED: ENABLE SCP AND RESTCONF ON ALL ONLINE DEVICES
# -------------------------------------------------------
def enable_scp_and_restconf(row, username, password):

    """
    PURPOSE
    -------
    Connects to a device via Netmiko and temporarily enables RESTCONF, SCP,
    and HTTPS (ip http secure-server) without saving to the startup config.
    Sets the device Status to ONLINE on success, OFFLINE on failure.
    If authentication fails, marks device as ONLINE with AUTH_BAD auth status.

    ARGUMENTS
    ---------
    row      (pd.Series): A single DataFrame row containing device information.
    username (str):       Device SSH username.
    password (str):       Device SSH password.

    RETURN VALUE
    ------------
    Tuple of (index, status, auth_status) where:
        status      : 'ONLINE' or 'OFFLINE'
        auth_status : 'AUTH_OK', 'AUTH_BAD', or None
    """

    # Define the parameters
    ip = row['OOBM IP Address']
    hostname = row['Hostname']
    index = row.name

    # Open the SSH connection
    try:
        connection = ConnectHandler(
            device_type='cisco_ios',
            host=ip,
            username=username,
            password=password
        )

        # Enable RESTCONF, SCP and HTTPS
        commands = [
            'ip http secure-server',
            'ip scp server enable',
            'restconf'
        ]

        # Push the configs to the devices
        connection.send_config_set(commands)
        connection.disconnect()

        #---------------------------------------------------------------
        logger.info(f"APIs enabled successfully on '{hostname}' ({ip})")
        #---------------------------------------------------------------

        # If it succeded, it means the device is ONLINE and the authentication succeded
        return index, 'ONLINE', 'AUTH_OK'
    
    except NetmikoAuthenticationException:
        #------------------------------------------------------------------------------------
        logger.warning(f"Authentication failed on '{hostname}' ({ip}) — device is reachable")
        #------------------------------------------------------------------------------------
        # Device responded but rejected credentials — it's ONLINE but AUTH_BAD
        return index, 'ONLINE', 'AUTH_BAD'
    
    except Exception as e:
        #--------------------------------------------------------------------------
        logger.warning(f"Device '{hostname}' ({ip}) is unreachable or failed: {e}")
        #--------------------------------------------------------------------------
        # If it failed, it means that the device is OFFLINE
        return index, 'OFFLINE', 'N/A'

# Thread pool -- Called by ...
def enable_scp_and_restconf_all(valid_devices_df, username, password):

    """
    PURPOSE
    -------
    Concurrently connects to all devices in the DataFrame and temporarily
    enables RESTCONF, SCP, and HTTPS.


    ARGUMENTS
    ---------
    valid_devices_df (pd.DataFrame): DataFrame containing at minimum 'OOBM IP Address' and 'Hostname' columns.
    username         (str):          Device SSH username.
    password         (str):          Device SSH password.


    RETURN VALUE
    ------------
    List of (index, status, auth_status) tuples — ONLINE/OFFLINE for status, AUTH_OK/AUTH_BAD/N/A for auth_status.
    """

    # Submit one configuration task per device — max 10 concurrent connections
    with ThreadPoolExecutor(max_workers=10) as executor:

        # Build a future for each device, keyed by hostname for result logging
        futures = {
            executor.submit(enable_scp_and_restconf, row, username, password): row['Hostname']
            for _, row in valid_devices_df.iterrows()
        }

        # Initialize the list that will collect (index, status) tuples
        results = []

        # Process results as each thread completes — order is not guaranteed
        for future in as_completed(futures):
            hostname = futures[future]
            try:
                # Unpack the (index, status, auth_status) tuple returned by enable_scp_and_restconf
                index, status, auth_status = future.result()
                results.append((index, status, auth_status))
                #-------------------------------------------------------------------------
                logger.info(f"'{hostname}' status set to: {status} | auth: {auth_status}")
                #-------------------------------------------------------------------------
            except Exception as e:
                #-------------------------------------------------------------
                logger.error(f"Unexpected error processing '{hostname}': {e}")
                #-------------------------------------------------------------

    return results


# NOT USED: DISABLE SCP AND HTTP ON ONLINE, NON-SELECTED DEVICES
# --------------------------------------------------------------
def disable_scp_and_restconf(row, username, password):

    """
    PURPOSE
    -------
    Connects to a device via Netmiko and disables RESTCONF, SCP, and HTTPS
    that were temporarily enabled earlier in the workflow. Does NOT save the
    running config – the changes are lost on reload (which is desired after
    a successful upgrade) but can be persisted elsewhere if needed.


    ARGUMENTS
    ---------
    row      (pd.Series): A single DataFrame row containing at minimum
                          'OOBM IP Address' and 'Hostname'.
    username (str):       Device SSH username.
    password (str):       Device SSH password.


    RETURN VALUE
    ------------
    Tuple of (index, result) where result is one of:
        'SUCCESS'        — commands were applied successfully.
        'CONNECT_ERROR'  — SSH connection or authentication failed.
        'UNEXPECTED_ERROR' — any other unhandled exception.
    """

    ip       = row['OOBM IP Address']
    hostname = row['Hostname']
    index    = row.name

    try:
        connection = ConnectHandler(
            device_type='cisco_ios',
            host=ip,
            username=username,
            password=password
        )

        # Disable the extra services that have been enabled
        commands = [
            'no ip scp server enable',
            'no restconf',
            'no ip http secure-server'
        ]
        #----------------------------------------------------------------
        logger.warning(f"Disabling HTTPS, RESTCONF and SCP on {ip}")
        #----------------------------------------------------------------
        connection.send_config_set(commands)

        #------------------------------------------------------------
        logger.warning(f"Disabled HTTPS, RESTCONF and SCP on {ip}")
        #------------------------------------------------------------

        connection.disconnect()
        return index, 'SUCCESS'

    except NetmikoAuthenticationException:
        #------------------------------------------------------------------------------------
        logger.error(f"'{hostname}' ({ip}): authentication failed – cannot disable services")
        #------------------------------------------------------------------------------------
        return index, 'CONNECT_ERROR'

    except Exception as e:
        #---------------------------------------------------------------------
        logger.error(f"'{hostname}' ({ip}): failed to disable services – {e}")
        #---------------------------------------------------------------------
        return index, 'UNEXPECTED_ERROR'

# Thread pool -- Called by ...
def disable_scp_and_restconf_all(online_authok_devices_df, username, password):
    
    """
    PURPOSE
    -------
    Concurrently disables RESTCONF, SCP, and HTTPS on all devices that are
    online and authenticated, using a thread pool. Returns a dictionary of
    results that can be used to populate the corresponding column in the
    DataFrame.


    ARGUMENTS
    ---------
    online_authok_devices_df (pd.DataFrame): DataFrame containing only devices
                                            that are ONLINE and AUTH_OK.
    username                 (str):          Device SSH username.
    password                 (str):          Device SSH password.


    RETURN VALUE
    ------------
    Dict of {index: result} — one entry per device, where index is the original
    DataFrame index and result is the return code from disable_scp_and_restconf.
    Devices that raised an unexpected thread exception are mapped to
    'UNEXPECTED_ERROR'.
    """

    results = {}

    with ThreadPoolExecutor(max_workers=10) as executor:

        # Keyed by dict to avoid tuple unpacking ambiguity on the DataFrame index
        futures = {
            executor.submit(disable_scp_and_restconf, row, username, password):
                {'index': idx, 'hostname': row['Hostname']}
            for idx, row in online_authok_devices_df.iterrows()
        }

        for future in as_completed(futures):
            idx      = futures[future]['index']
            hostname = futures[future]['hostname']
            try:
                index, result = future.result()
                #---------------------------------------------------------
                logger.info(f"'{hostname}': disable services result — {result}")
                #---------------------------------------------------------
            except Exception as e:
                #-----------------------------------------------------------
                logger.error(f"'{hostname}': unexpected thread error — {e}")
                #-----------------------------------------------------------
                index  = idx
                result = 'UNEXPECTED_ERROR'

            results[index] = result

    return results


# CHECK SCP SERVER STATUS ON ONLINE/AUTH_OK DEVICES
# -------------------------------------------------
def check_scp_status(row, username, password):

    """
    PURPOSE
    -------
    Connects to a device via Netmiko and runs 'show run | section ip scp'
    to determine whether the SCP server is enabled on the device.

    The command returns an empty string when SCP is disabled and the
    line 'ip scp server enable' when it is enabled. No configuration
    change is made — this is a read-only check.


    ARGUMENTS
    ---------
    row      (pd.Series): A single DataFrame row containing 'OOBM IP Address' and 'Hostname'.
    username (str):       Device SSH username.
    password (str):       Device SSH password.


    RETURN VALUE
    ------------
    Tuple of (index, scp_status) where scp_status is one of:
        'YES'  — 'ip scp server enable' found in running config.
        'NO' — Command returned empty output; SCP is not configured.
        'AUTH_BAD' — Device is reachable but rejected the credentials.
        'UNKNOWN'  — SSH connection could not be established.
    """

    ip       = row['OOBM IP Address']
    hostname = row['Hostname']
    index    = row.name

    try:
        connection = ConnectHandler(
            device_type='cisco_ios',
            host=ip,
            username=username,
            password=password
        )

        output = connection.send_command('show run | section ip scp')
        connection.disconnect()

        # Determine if SCP is enabled or not based on the output
        scp_enabled = 'ip scp server enable' in output.lower()
        status = 'YES' if scp_enabled else 'NO'

        #---------------------------------------------------------
        logger.info(f"SCP check on '{hostname}' ({ip}): {status}")
        #---------------------------------------------------------

        return index, status

    except NetmikoAuthenticationException:
        #-------------------------------------------------------------------------------
        logger.warning(f"Authentication failed on '{hostname}' ({ip}) during SCP check")
        #-------------------------------------------------------------------------------
        return index, 'AUTH_BAD'

    except Exception as e:
        #------------------------------------------------------------------------
        logger.warning(f"Could not reach '{hostname}' ({ip}) for SCP check: {e}")
        #------------------------------------------------------------------------
        return index, 'UNKNOWN'

# Thread pool -- Called by '' from 'excel_and_data_ops'   
def check_scp_status_all(online_authok_df, username, password):

    """
    PURPOSE
    -------
    Concurrently checks SCP server status on all devices in the supplied
    DataFrame using a thread pool. Each device is handled by check_scp_status.


    ARGUMENTS
    ---------
    online_authok_df (pd.DataFrame): DataFrame pre-filtered to ONLINE/AUTH_OK devices.
                                     Must contain 'OOBM IP Address' and 'Hostname'.
    username         (str):          Device SSH username.
    password         (str):          Device SSH password.


    RETURN VALUE
    ------------
    List of (index, scp_status) tuples — one per device.
    index matches the original DataFrame index for safe .at[] assignment.
    scp_status is one of: 'YES', 'NO', 'AUTH_BAD', 'UNKNOWN'.
    """

    results = []

    with ThreadPoolExecutor(max_workers=10) as executor:

        futures = {
            executor.submit(check_scp_status, row, username, password): row['Hostname']
            for _, row in online_authok_df.iterrows()
        }

        for future in as_completed(futures):
            hostname = futures[future]
            try:
                index, scp_status = future.result()
                results.append((index, scp_status))
                #----------------------------------------------------
                logger.info(f"'{hostname}' SCP status: {scp_status}")
                #----------------------------------------------------
            except Exception as e:
                #------------------------------------------------------------------
                logger.error(f"Unexpected error checking SCP on '{hostname}': {e}")
                #------------------------------------------------------------------

    return results


# GET THE DEVICE'S FLASH MEMORY FREE SPACE
# ----------------------------------------
def get_flash_free_space(row, username, password):

    """
    PURPOSE
    -------
    Connects to a device using Netmiko and retrieves free space (in Bytes)
    from the primary filesystem (bootflash:/flash:) by parsing CLI output.

    Designed to match previous asyncio-style return format.

    
    ARGUMENTS
    ---------
    row      (pd.Series): A single DataFrame row containing device information.
    username (str):       Device SSH username.
    password (str):       Device SSH password.

    
    RETURN VALUE
    ------------
    Tuple of (ip, free_bytes, error_code) where:
        free_bytes : Free space in Bytes (int) or None if unavailable
        error_code : None on success, or string indicating failure type:
            - "NO_FLASH_FOUND"      : No flash filesystem found in output
            - "PARSE_ERROR"         : Could not parse free bytes from output
            - "UNEXPECTED_ERROR"    : Any other exception
    """

    ip = row['OOBM IP Address']
    hostname = row['Hostname']

    try:
        connection = ConnectHandler(
            device_type='cisco_ios',
            host=ip,
            username=username,
            password=password
        )

        output = connection.send_command("show file systems")
        connection.disconnect()

        # Initialize the empty list of lines that may contain flash memory information
        flash_candidates = []

        # Find flash/bootflash filesystem lines
        for line in output.splitlines():
            line_clean = line.strip().lower()

            if (
                ("bootflash:" in line_clean or "flash:" in line_clean) and
                any(char.isdigit() for char in line_clean)
            ):
                flash_candidates.append(line)

        if not flash_candidates:
            #-------------------------------------------------------------------
            logger.warning(f"No flash filesystem found for '{hostname}' ({ip})")
            #-------------------------------------------------------------------
            return ip, None, "NO_FLASH_FOUND"

        # Intialize the empty list of lines that actually contain the desired information
        selected_line = None

        # Prefer bootflash if present
        for line in flash_candidates:
            if "bootflash:" in line.lower():
                selected_line = line
                break

        # Otherwise use first flash match
        if not selected_line:
            selected_line = flash_candidates[0]

        # Split and keep only numeric tokens — skips '*', 'rw', 'flash:', etc.
        tokens = [t for t in selected_line.split() if t.isdigit()]

        # Expect at least 3 tokens: total_bytes, free_bytes, and at minimum one more field
        if len(tokens) < 2:
            #--------------------------------------------------------------------------------------------------------------
            logger.warning(f"Invalid filesystem data for '{hostname}' ({ip}): unexpected format '{selected_line.strip()}'")
            #--------------------------------------------------------------------------------------------------------------
            return ip, None, "PARSE_ERROR"

        try:
            # tokens[0] = total size, tokens[1] = free size
            free_bytes = int(tokens[1])
        except ValueError:
            #---------------------------------------------------------------------------------------------
            logger.warning(f"Could not parse free bytes for '{hostname}' ({ip}): token was '{tokens[1]}'")
            #---------------------------------------------------------------------------------------------
            return ip, None, "PARSE_ERROR"

        #-------------------------------------------------------------------------------------------------
        logger.info(f"Flash free space for {ip}: {free_bytes} Bytes -- {free_bytes / (1024*1024):.0f} MB")
        #-------------------------------------------------------------------------------------------------

        # Explicitely return a 'None' error code, so that it does not cause a crash on the 'excel_and_data_ops.populate_flash_free_space_column' function,
        # since is expecting to recieve a tuple of 3 values.
        return ip, free_bytes, None

    except Exception as e:
        #------------------------------------------------------------------------
        logger.warning(f"Failed to get flash space for '{hostname}' ({ip}): {e}")
        #------------------------------------------------------------------------
        return ip, None, 'UNEXPECTED_ERROR'
    
# Thread pool -- Called by 'populate_flash_free_space_column' from 'excel_and_data_ops'
def get_flash_free_space_all(valid_devices_df, username, password):
    """
    PURPOSE
    -------
    Executes get_flash_free_space() across all devices using multithreading.

    ARGUMENTS
    ---------
    valid_devices_df (pd.DataFrame): Device inventory
    username         (str):          SSH username
    password         (str):          SSH password

    RETURN VALUE
    ------------
    List of tuples:
        (ip, free_bytes)
    """

    results = []

    with ThreadPoolExecutor(max_workers=10) as executor:

        # Keyed by hostname for result logging — mirrors enable_scp_and_restconf_all pattern
        futures = {
            executor.submit(get_flash_free_space, row, username, password): row['Hostname']
            for _, row in valid_devices_df.iterrows()
        }

        for future in as_completed(futures):
            hostname = futures[future]
            try:
                results.append(future.result())
            except Exception as e:
                #--------------------------------------------------------------------------
                logger.error(f"Unexpected error getting flash space for '{hostname}': {e}")
                #--------------------------------------------------------------------------

    return results


# PUSH COMMANDS FOR BOOT VARIABLE AND INSTALL UPGRADE
# ---------------------------------------------------
def install_ios_image(row, username, password):

    """
    PURPOSE
    -------
    Connects to a device via Netmiko and executes the IOS-XE install mode
    upgrade sequence in a single SSH session:

        1. Pushes boot system commands and saves the config.

        2. Runs 'install add file bootflash:<bin>' — stages the image on bootflash.
           Waits for SUCCESS, ERROR, or FAILED in the output to confirm completion.

        3. Runs 'install activate' via send_multiline — waits for the reload
           confirmation prompt, auto-answers 'y', and waits for 'Starting Activate'
           to confirm the reload sequence has begun.

    User confirmation is handled upstream by confirm_installs() before this
    function is ever called. Version verification and install commit after the
    reboot are handled separately by the caller via the post-install pipeline.


    ARGUMENTS
    ---------
    row      (pd.Series): A single DataFrame row containing at minimum 'OOBM IP Address', 'Hostname', and 'IOS Image Path' columns.
    username (str):       Device SSH username.
    password (str):       Device SSH password.

    DEAD DEVICE DETECTION
    ---------------------
    The install add command can legitimately take up to 15 minutes, so read_timeout
    is set to 900s. However, if the device goes offline mid-command, Paramiko returns
    empty bytes from recv() rather than raising — causing Netmiko to silently loop
    until read_timeout expires.

    Two mechanisms are combined to detect dead devices within ~70 seconds instead:

        TCP keepalives (OS level):
            Configured on the raw socket under Paramiko. After 60s of idle silence
            the OS sends 3 probe packets at 10s intervals. If all fail, the OS marks
            the socket as broken (~90s total). On the next recv() Netmiko gets an
            OSError, which is caught as CONNECT_ERROR.

        Watchdog thread:
            A daemon thread polls transport.is_active() every 10 seconds. The moment
            the transport goes inactive (which happens as soon as TCP keepalives fail),
            the watchdog force-closes the raw socket. This converts Paramiko's silent
            empty-read loop into an immediate OSError, breaking out of send_command
            within one 10-second poll interval rather than waiting for read_timeout.

    The watchdog is always stopped via a threading.Event in the finally block,
    regardless of whether the function succeeds, fails, or raises an exception.

    RETURN VALUE
    ------------
    Tuple of (index, result) where result is one of:

        'INSTALL_TRIGGERED'  — activate confirmed, device is rebooting.
        'ADD_FAILED'         — install add did not complete successfully.
        'ACTIVATE_FAILED'    — install activate returned a FAILED string.
        'CONNECT_ERROR'      — SSH connection failed, authentication rejected,
                               or dead peer detected by TCP keepalive/watchdog.
        'UNEXPECTED_ERROR'   — uncaught exception during any stage.
    """

    index    = row.name
    ip       = row['OOBM IP Address']
    hostname = row['Hostname']
    bin_file = os.path.basename(row['IOS Image Path'])

    # Initialize the watchdog
    _stop_watchdog = None

    try:
        connection = ConnectHandler(
            device_type='cisco_ios',
            host=ip,
            username=username,
            password=password,
            session_timeout=900,
            keepalive=60
        )

        ### <=== TCP KEEPALIVE + WATCHDOG THREAD ===> ###
        # TCP keepalives let the OS detect a dead peer within ~40s (60s idle + 3x10s probes).
        # However, when the socket dies, Paramiko returns b"" from recv() rather than raising —
        # so Netmiko keeps looping silently until read_timeout expires (15 min).
        # The watchdog thread polls transport.is_active() every 10 seconds and force-closes
        # the socket the moment it goes False, which raises OSError inside Netmiko's recv()
        # loop and breaks out immediately.

        # Netmiko uses Paramiko under the hood (a Python SSH library). Paramiko wraps a normal TCP socket for the SSH connection. That TCP socket is managed 
        # by your operating system. When we set sock.setsockopt(socket.SOL_SOCKET, socket.SO_KEEPALIVE, 1), we are telling the operating system:

        ##Hey OS, please monitor this TCP connection for me. If it goes silent for too long, send probe packets to check if the remote side is still alive.##

        # The key thing is: Netmiko and Paramiko don't send the keepalive probes themselves — the OS kernel does it in the background, independently of Python. Netmiko 
        # doesn't even know it's happening.

        # So the chain is:

        #     connection.remote_conn is the Paramiko Channel object.

        #     .get_transport() gets the underlying Transport object.

        #     .sock is the raw Python socket.socket object.

        # We call setsockopt() on it to configure the OS-level TCP keepalive. The OS starts the keepalive timer (after TCP_KEEPIDLE idle seconds). If probes fail, the OS 
        # marks the socket as broken. Next time Netmiko tries to recv() data (during send_command), it gets a socket error and raises an exception.

        # When we call setsockopt() on that .sock, we are instructing the operating system to apply the TCP keepalive settings to exactly the TCP 
        # connection that carries this device's SSH session. No other connections are affected, and we aren't guessing — we're directly accessing 
        # the socket stored inside the Paramiko Transport that Netmiko is using.
        try:
            transport = connection.remote_conn.get_transport()
            sock = transport.sock
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_KEEPALIVE, 1)

            if sys.platform == 'win32':
                sock.ioctl(0x98000004, (1, 60000, 10000))  # idle=60s, intvl=10s
                #-----------------------------------------------------------------------
                logger.debug(f"'{hostname}' ({ip}): TCP keepalive configured (Windows)")
                #-----------------------------------------------------------------------
            else:
                if hasattr(socket, 'TCP_KEEPIDLE'):
                    sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_KEEPIDLE,  60)
                if hasattr(socket, 'TCP_KEEPINTVL'):
                    sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_KEEPINTVL, 10)
                if hasattr(socket, 'TCP_KEEPCNT'):
                    sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_KEEPCNT,    3)
                #---------------------------------------------------------------------
                logger.debug(f"'{hostname}' ({ip}): TCP keepalive configured (Linux)")
                #---------------------------------------------------------------------

        except Exception as keepalive_err:
            #------------------------------------------------------------------------------------
            logger.warning(f"'{hostname}' ({ip}): could not set TCP keepalive — {keepalive_err}")
            #------------------------------------------------------------------------------------
            transport = None
            sock = None


        # It doesn't matter where we place it within the function — the watchdog runs on a separate daemon thread that monitors independently of whatever 
        # the main thread is doing. Once _watchdog.start() is called, it keeps polling transport.is_active() every 10 seconds continuously until either 
        # _stop_watchdog.set() fires in the finally block, or it detects a dead transport and breaks itself. So whether the main thread is currently inside send_config_set, 
        # save_config, send_command (install add), or send_multiline (install activate) — the watchdog is watching the entire time.


        # Start the watchdog — it fires every 10s and closes the socket if the transport dies
        _stop_watchdog = threading.Event()

        def _transport_watchdog():
            while not _stop_watchdog.wait(timeout=10):
                try:
                    if transport and not transport.is_active():
                        #--------------------------------------------------------------------------------
                        logger.warning(f"'{hostname}' ({ip}): transport inactive — force-closing socket")
                        #--------------------------------------------------------------------------------

                        # Close the Paramiko channel Netmiko is blocking on.
                        # This raises SSHException / EOFError in the recv() loop immediately.
                        try:
                            connection.remote_conn.close()
                        except Exception:
                            pass

                        # Shutdown the socket at OS level BEFORE closing it.
                        # shutdown(SHUT_RDWR) is the critical step — it immediately unblocks
                        # any recv() sitting in the kernel, raising an OSError right away.
                        # close() alone only marks the fd for cleanup; the blocked recv()
                        # keeps waiting until the OS-level timeout (which is what causes the
                        # 15-minute hang).
                        if sock:
                            try:
                                sock.shutdown(socket.SHUT_RDWR)
                                #-------------------------------------------------------
                                logger.warning(f"'{hostname}' ({ip}): socket shut down")
                                #-------------------------------------------------------
                            except Exception:
                                #--------------------------------------------------------------------------------------------
                                logger.warning(f"'{hostname}' ({ip}): failed to shutting down socket, passing to closing it")
                                #--------------------------------------------------------------------------------------------
                                pass
                            try:
                                sock.close()
                                #----------------------------------------------------
                                logger.warning(f"'{hostname}' ({ip}): socket closed")
                                #----------------------------------------------------
                            except Exception:
                                pass

                        # Tear down the Paramiko transport last.
                        try:
                            transport.close()
                            #----------------------------------------------------------------
                            logger.warning(f"'{hostname}' ({ip}): Paramiko transport closed")
                            #----------------------------------------------------------------
                        except Exception:
                            pass
                        
                        break

                except Exception:
                    break

        _watchdog = threading.Thread(target=_transport_watchdog, daemon=True)
        _watchdog.start()


        ### <=== PUSH BOOT SYSTEM COMMANDS AND SAVE ===> ###
        # Set the boot variable and save before sending the install commands —
        # this ensures the device boots from the correct image even if the
        # session is interrupted after the reload fires.
        boot_commands = [
            'no boot system',
            'boot system bootflash:packages.conf',
        ]

        #----------------------------------------------------------------
        logger.info(f"'{hostname}' ({ip}): pushing boot system commands")
        #----------------------------------------------------------------

        connection.send_config_set(boot_commands)
        connection.save_config()

        #--------------------------------------------------------------------------------
        logger.info(f"'{hostname}' ({ip}): boot system commands pushed and config saved")
        #--------------------------------------------------------------------------------


        ### <=== INSTALL ADD ===> ###
        # Stages the image on bootflash. This can take several minutes depending
        # on image size — read_timeout is set generously to avoid premature timeout.
        #------------------------------------------------------------------------
        logger.info(f"'{hostname}' ({ip}): running install add for '{bin_file}'")
        #------------------------------------------------------------------------

        add_output = connection.send_command(
            f'install add file bootflash:{bin_file}',
            expect_string=r'SUCCESS|ERROR|FAILED',
            read_timeout=900,
            strip_prompt=False,
            strip_command=False
        )

        if 'SUCCESS' not in add_output:
            #------------------------------------------------------------------------------
            logger.error(f"'{hostname}' ({ip}): install add failed — output: {add_output}")
            #------------------------------------------------------------------------------
            return index, 'ADD_FAILED'
        
        else:
            #------------------------------------------------------------------------
            logger.info(f"'{hostname}' ({ip}): install add succeeded - {add_output}")
            #------------------------------------------------------------------------

            time.sleep(5)

            ### <=== INSTALL ACTIVATE ===> ###
            #------------------------------------------------------------
            logger.info(f"'{hostname}' ({ip}): running install activate")
            #------------------------------------------------------------


            # Wait for the [y/n] confirmation prompt before answering —
            # send_command blocks until the prompt appears
            
            # Command list. This is a list of lists with each element of the inner list being the command to send and the pattern to search for.
            cmd_list = [
                    ['install activate', r"This operation may require a reload"],
                    #  Use a null-string to indicate that we want to automatically search for the trailing prompt as the pattern.
                    ['y', r"Starting Activate"],
                ]
            
            activate_output = connection.send_multiline(
                cmd_list,
                read_timeout=120
            )
            # activate_output = connection.send_command(
            #     'install activate',
            #     expect_string='[y/n]',
            #     read_timeout=120
            # )

            if 'FAILED:' in activate_output:
                #----------------------------------------------------------------------------------------
                logger.error(f"'{hostname}' ({ip}): install activate failed — output: {activate_output}")
                #----------------------------------------------------------------------------------------
                return index,'ACTIVATE_FAILED'

            else:
                #-------------------------------------------------------------------------------------------------
                logger.info(f"'{hostname}' ({ip}): install activate sent — device reloading -- {activate_output}")
                #-------------------------------------------------------------------------------------------------
                connection.disconnect()
                return index, 'INSTALL_TRIGGERED'
                            
    except NetmikoAuthenticationException:
        #---------------------------------------------------------------------------
        logger.error(f"'{hostname}' ({ip}): authentication failed — cannot proceed")
        #---------------------------------------------------------------------------
        return index, 'CONNECT_ERROR'

    except (OSError, EOFError, NetmikoTimeoutException, SSHException) as e:
        # OSError:                  watchdog force-closed the TCP socket
        # SSHException:             watchdog called transport.close()
        # EOFError/Timeout:         other connection loss paths
        #-------------------------------------------------------------------------
        logger.error(f"'{hostname}' ({ip}): connection lost during install — {e}")
        #-------------------------------------------------------------------------
        return index, 'CONNECT_ERROR'

    except Exception as e:
        #-------------------------------------------------------------------------
        logger.error(f"'{hostname}' ({ip}): failed during install sequence — {e}")
        #-------------------------------------------------------------------------
        return index, 'UNEXPECTED_ERROR'
    
    finally:
        if _stop_watchdog is not None:
            _stop_watchdog.set()  # Stop the watchdog thread regardless of outcome

# Thread pool -- Called by 'populate_install_status_column' from 'excel_and_data_ops'
def install_ios_image_all(confirmed_df, username, password):

    """
    PURPOSE
    -------
    Concurrently executes the IOS install sequence across all confirmed devices
    using a thread pool. User confirmation has already been collected upstream
    by confirm_installs() and all devices are guaranteed to have Transfer Status
    SUCCESS — filtered upstream by get_install_eligible_devices_df().


    ARGUMENTS
    ---------
    confirmed_df (pd.DataFrame): DataFrame containing only devices the user
                                 confirmed for install — subset of install_eligible_devices_df.
    username     (str):          Device SSH username.
    password     (str):          Device SSH password.


    RETURN VALUE
    ------------
    Dict of {index: result} — one entry per device, where index is the original
    DataFrame index and result is the return code from install_ios_image.
    Devices that raised an unexpected thread exception are mapped to 'UNEXPECTED_ERROR'.
    """

    results = {}

    with ThreadPoolExecutor(max_workers=10) as executor:

        # Keyed by dict to avoid tuple unpacking ambiguity on the DataFrame index
        futures = {
            executor.submit(install_ios_image, row, username, password): {'index': idx, 'hostname': row['Hostname']}
            for idx, row in confirmed_df.iterrows()
        }

        # Process results as each thread completes — order is not guaranteed
        for future in as_completed(futures):
            idx      = futures[future]['index']
            hostname = futures[future]['hostname']
            try:
                index, result = future.result()
                #------------------------------------------------------
                logger.info(f"'{hostname}': install result — {result}")
                #------------------------------------------------------
            except Exception as e:
                #-----------------------------------------------------------
                logger.error(f"'{hostname}': unexpected thread error — {e}")
                #-----------------------------------------------------------
                index  = idx
                result = 'UNEXPECTED_ERROR'

            results[index] = result

    return results


# PUSH INSTALL COMMIT
# -------------------
def commit_ios_install(row, username, password):

    """
    PURPOSE
    -------
    Connects to a device via Netmiko and runs 'install commit' to persist
    the activated IOS version as the new default boot image. Called after
    the device has come back up following install activate.


    ARGUMENTS
    ---------
    row      (pd.Series): A single DataFrame row containing device information.
    username (str):       Device SSH username.
    password (str):       Device SSH password.


    RETURN VALUE
    ------------
    Tuple of (index, result) where result is one of:
        'COMMIT_SUCCESS'   — install commit completed successfully.
        'COMMIT_FAILED'    — install commit returned an error or unexpected output.
        'CONNECT_ERROR'    — SSH connection failed.
        'UNEXPECTED_ERROR' — uncaught exception.
    """

    index    = row.name
    ip       = row['OOBM IP Address']
    hostname = row['Hostname']

    try:
        connection = ConnectHandler(
            device_type='cisco_ios',
            host=ip,
            username=username,
            password=password
        )

        #------------------------------------------------------------------
        logger.info(f"'{hostname}' ({ip}): running install commit")
        #------------------------------------------------------------------

        commit_output = connection.send_command(
            'install commit',
            expect_string=r'SUCCESS|ERROR|FAILED',
            read_timeout=300
        )

        if 'SUCCESS' not in commit_output:
            #------------------------------------------------------------------------------------
            logger.error(f"'{hostname}' ({ip}): install commit failed — output: {commit_output}")
            #------------------------------------------------------------------------------------
            return index, 'COMMIT_FAILED'

        #------------------------------------------------------------------
        logger.info(f"'{hostname}' ({ip}): install commit succeeded")
        #------------------------------------------------------------------
        return index, 'COMMIT_SUCCESS'

    except NetmikoAuthenticationException:
        #---------------------------------------------------------------------------
        logger.error(f"'{hostname}' ({ip}): authentication failed during commit")
        #---------------------------------------------------------------------------
        return index, 'CONNECT_ERROR'

    except Exception as e:
        #---------------------------------------------------------------------------
        logger.error(f"'{hostname}' ({ip}): unexpected error during commit — {e}")
        #---------------------------------------------------------------------------
        return index, 'UNEXPECTED_ERROR'

    finally:
        try:
            connection.disconnect()
        except Exception:
            pass

# Thread pool -- Called by 'populate_post_install_status_column' from 'excel_and_data_ops'
def commit_ios_install_all(online_devices_df, username, password):

    """
    PURPOSE
    -------
    Concurrently runs 'install commit' across all devices that came back
    online after the install activate reload, using a thread pool.


    ARGUMENTS
    ---------
    online_devices_df (pd.DataFrame): DataFrame containing only devices that
                                      came back online after the reload.
    username          (str):          Device SSH username.
    password          (str):          Device SSH password.


    RETURN VALUE
    ------------
    Dict of {index: result} — one entry per device, where index is the original
    DataFrame index and result is the return code from commit_ios_install.
    Devices that raised an unexpected thread exception are mapped to 'UNEXPECTED_ERROR'.
    """

    results = {}

    with ThreadPoolExecutor(max_workers=10) as executor:

        # Keyed by dict to avoid tuple unpacking ambiguity on the DataFrame index
        futures = {
            executor.submit(commit_ios_install, row, username, password): {'index': idx, 'hostname': row['Hostname']}
            for idx, row in online_devices_df.iterrows()
        }

        # Process results as each thread completes — order is not guaranteed
        for future in as_completed(futures):
            idx      = futures[future]['index']
            hostname = futures[future]['hostname']
            try:
                index, result = future.result()
                #------------------------------------------------------
                logger.info(f"'{hostname}': commit result — {result}")
                #------------------------------------------------------
            except Exception as e:
                #-----------------------------------------------------------
                logger.error(f"'{hostname}': unexpected thread error — {e}")
                #-----------------------------------------------------------
                index  = idx
                result = 'UNEXPECTED_ERROR'

            results[index] = result

    return results


# REMOVE INACTIVE IOS PACKAGES FROM DEVICES THAT THE INSTALL ADD SUCCEDED
# -----------------------------------------------------------------------
def remove_inactive_ios(row, username, password):

    """
    PURPOSE
    -------
    Connects to a device via Netmiko and runs 'install remove inactive' to
    clean up inactive IOS package files from bootflash, auto-answering the
    confirmation prompt. Only intended to be called on devices whose Install
    Status is INSTALL_TRIGGERED — enforced upstream by the caller.


    ARGUMENTS
    ---------
    row      (pd.Series): A single DataFrame row containing device information.
    username (str):       Device SSH username.
    password (str):       Device SSH password.


    RETURN VALUE
    ------------
    Tuple of (index, result) where result is one of:
        'CLEANED'          — inactive packages removed successfully.
        'CLEAN_FAILED'     — command returned an error or unexpected output.
        'NOTHING_TO_CLEAN' — no inactive packages found, nothing was removed.
        'CONNECT_ERROR'    — SSH connection failed.
        'UNEXPECTED_ERROR' — uncaught exception.
    """

    index    = row.name
    ip       = row['OOBM IP Address']
    hostname = row['Hostname']

    try:
        connection = ConnectHandler(
            device_type='cisco_ios',
            host=ip,
            username=username,
            password=password
        )

        #-------------------------------------------------------------------
        logger.info(f"'{hostname}' ({ip}): running install remove inactive")
        #-------------------------------------------------------------------

        # Command list. This is a list of lists with each element of the inner list being the command to send and the pattern to search for.
        cmd_list = [
                ['install remove inactive', r"Do you want to remove the above files"],
                #  Use a null-string to indicate that we want to automatically search for the trailing prompt as the pattern.
                ['y', r"Deleting file"]
            ]
        
        remove_output = connection.send_multiline(
            cmd_list,
            read_timeout=120
        )

        # If there is nothing to remove the device skips the prompt entirely
        # and returns immediately — detect this before trying to answer
        if 'No inactive' in remove_output or 'Nothing to clean' in remove_output:
            #------------------------------------------------------------------
            logger.info(f"'{hostname}' ({ip}): no inactive packages to remove")
            #------------------------------------------------------------------
            return index, 'NOTHING_TO_CLEAN'

        if 'SUCCESS' not in remove_output:
            #---------------------------------------------------------------------------------------------
            logger.error(f"'{hostname}' ({ip}): install remove inactive failed — output: {remove_output}")
            #---------------------------------------------------------------------------------------------
            return index, 'CLEAN_FAILED'

        #--------------------------------------------------------------------------
        logger.info(f"'{hostname}' ({ip}): inactive packages removed successfully")
        #--------------------------------------------------------------------------
        return index, 'CLEANED'

    except NetmikoAuthenticationException:
        #-------------------------------------------------------------------------
        logger.error(f"'{hostname}' ({ip}): authentication failed during cleanup")
        #-------------------------------------------------------------------------
        return index, 'CONNECT_ERROR'

    except Exception as e:
        #--------------------------------------------------------------------------
        logger.error(f"'{hostname}' ({ip}): unexpected error during cleanup — {e}")
        #--------------------------------------------------------------------------
        return index, 'UNEXPECTED_ERROR'

    finally:
        try:
            connection.disconnect()
        except Exception:
            pass

def remove_inactive_ios_all(valid_devices_df, username, password):

    """
    PURPOSE
    -------
    Concurrently runs 'install remove inactive' across all devices whose
    Install Status is INSTALL_TRIGGERED, using a thread pool.


    ARGUMENTS
    ---------
    valid_devices_df (pd.DataFrame): Full device DataFrame — filtered internally
                                     to INSTALL_TRIGGERED devices only.
    username         (str):          Device SSH username.
    password         (str):          Device SSH password.


    RETURN VALUE
    ------------
    Dict of {index: result} — one entry per device, where index is the original
    DataFrame index and result is the return code from remove_inactive_ios.
    Devices that raised an unexpected thread exception are mapped to 'UNEXPECTED_ERROR'.
    """

    # Only run cleanup on devices that had the install triggered —
    # all other devices never received a new image so there is nothing to clean
    triggered_devices = valid_devices_df[
        valid_devices_df['Install Status'] == 'INSTALL_TRIGGERED'
    ]

    results = {}

    # Pre-populate N/A for all non-triggered devices upfront —
    # cleanup is not applicable for devices that never received an install
    results = {
        idx: 'N/A'
        for idx in valid_devices_df.index
        if idx not in triggered_devices.index
    }

    if triggered_devices.empty:
        #----------------------------------------------------------------------
        logger.warning("No INSTALL_TRIGGERED devices found — skipping cleanup")
        #----------------------------------------------------------------------
        return results

    with ThreadPoolExecutor(max_workers=10) as executor:

        # Keyed by dict to avoid tuple unpacking ambiguity on the DataFrame index
        futures = {
            executor.submit(remove_inactive_ios, row, username, password): {'index': idx, 'hostname': row['Hostname']}
            for idx, row in triggered_devices.iterrows()
        }

        # Process results as each thread completes — order is not guaranteed
        for future in as_completed(futures):
            idx      = futures[future]['index']
            hostname = futures[future]['hostname']
            try:
                index, result = future.result()
                #------------------------------------------------------
                logger.info(f"'{hostname}': cleanup result — {result}")
                #------------------------------------------------------
            except Exception as e:
                #-----------------------------------------------------------
                logger.error(f"'{hostname}': unexpected thread error — {e}")
                #-----------------------------------------------------------
                index  = idx
                result = 'UNEXPECTED_ERROR'

            results[index] = result

    return results


# CANCEL ALL ACTIVE TRANSFERS
# ---------------------------
def cancel_single_device(row, username, password):

    """
    PURPOSE
    -------
    Connect to one device via Netmiko, clear all VTY lines to terminate any
    active SCP transfer, then reconnect and delete the partially transferred
    file from bootflash.

    The global cancel_event must already be set before this function is called
    (done upstream in the thread‑pool wrapper) to prevent new transfers from
    starting while the cancellation is in progress.


    ARGUMENTS
    ---------
    row      (pd.Series): A single DataFrame row containing at minimum 'OOBM IP Address', 'Hostname', and 'IOS Image Path'.
    username (str):       Device SSH username.
    password (str):       Device SSH password.


    RETURN VALUE
    ------------
    Tuple of (index, result) where result is one of:
        'SUCCESS'           — VTY lines cleared and file deleted successfully.
        'VTY_CLEAR_FAILED'  — initial connection or VTY clearing failed.
        'FILE_DELETE_FAILED'— VTY cleared but file deletion failed.
        'UNEXPECTED_ERROR'  — uncaught exception during the attempt.
    """

    ip       = row['OOBM IP Address']
    hostname = row['Hostname']
    index    = row.name
    remote_filename = os.path.basename(row['IOS Image Path'])

    try:
        # First connection: clear VTY lines
        connection = ConnectHandler(
            device_type='cisco_ios',
            host=ip,
            username=username,
            password=password
        )

        #----------------------------------------------------
        logger.warning(f"Sending 'clear line vty *' to {ip}")
        #----------------------------------------------------
        for vty_line in range(0, 16):
            connection.send_command(f'clear line vty {vty_line}',
                                    expect_string=r'#|\[confirm\]')
            connection.send_command('\n', expect_string=r'#')
        connection.disconnect()
        #-----------------------------------------------------
        logger.warning(f"Active VTY sessions cleared on {ip}")
        #-----------------------------------------------------

    except Exception as e:
        #-----------------------------------------------------------
        logger.error(f"'{hostname}' ({ip}): VTY clear failed – {e}")
        #-----------------------------------------------------------
        return index, 'VTY_CLEAR_FAILED'

    # Only attempt file deletion if VTY clearing succeeded
    try:
        connection = ConnectHandler(
            device_type='cisco_ios',
            host=ip,
            username=username,
            password=password
        )

        # Delete the partially copied IOS image file
        connection.send_command(f'delete /force bootflash:{remote_filename}')
        #-------------------------------------------------------------
        logger.warning(f"Partial file deleted from bootflash on {ip}")
        #-------------------------------------------------------------

        # Close the connection
        connection.disconnect()
        
        return index, 'SUCCESS'

    except Exception as e:
        #---------------------------------------------------------------
        logger.error(f"'{hostname}' ({ip}): file deletion failed – {e}")
        #---------------------------------------------------------------
        return index, 'FILE_DELETE_FAILED'

    return index, 'VTY_CLEAR_FAILED' # fallback

# Thread pool -- Called by 'populate_post_install_columns' from 'excel_and_data_ops'
def cancel_active_transfers_all(selected_devices_df, username, password):

    """
    PURPOSE
    -------
    Set the global cancel_event to prevent new SCP transfers from starting,
    then concurrently clear VTY lines and delete partial files on all selected
    devices using a thread pool.


    ARGUMENTS
    ---------
    selected_devices_df (pd.DataFrame): DataFrame containing the devices with
                                        active transfers.
    username            (str):          Device SSH username.
    password            (str):          Device SSH password.


    RETURN VALUE
    ------------
    Dict of {index: result} — one entry per device, where index is the original
    DataFrame index and result is the return code from cancel_single_device.
    Devices that raised an unexpected thread exception are mapped to
    'UNEXPECTED_ERROR'.
    """

    # Prevent any new transfers from starting
    cancel_event.set()
    #---------------------------------------------------------------------------------
    logger.warning("Cancellation requested — clearing VTY sessions on all devices...")
    #---------------------------------------------------------------------------------

    results = {}
    with ThreadPoolExecutor(max_workers=len(selected_devices_df)) as executor:
        futures = {
            executor.submit(cancel_single_device, row, username, password):
                {'index': idx, 'hostname': row['Hostname']}
            for idx, row in selected_devices_df.iterrows()
        }

        for future in as_completed(futures):
            idx      = futures[future]['index']
            hostname = futures[future]['hostname']
            try:
                index, result = future.result()
                #-----------------------------------------------------
                logger.info(f"'{hostname}': cancel result — {result}")
                #-----------------------------------------------------
            except Exception as e:
                #-----------------------------------------------------------
                logger.error(f"'{hostname}': unexpected thread error — {e}")
                #-----------------------------------------------------------

                index  = idx
                result = 'UNEXPECTED_ERROR'
            results[index] = result

    return results


# PUSH EXEC COMMANDS TO A SINGLE DEVICE
# --------------------------------------
def push_commands(row, username, password, commands):

    """
    PURPOSE
    -------
    Opens a Netmiko SSH connection to a single device and sends each command
    in the list sequentially using send_command. Commands are treated as
    exec-mode only — no config mode is entered. If the user needs config-mode
    commands, they must include 'configure terminal' as the first line.


    ARGUMENTS
    ---------
    row      (pd.Series): A single DataFrame row containing 'OOBM IP Address' and 'Hostname'.
    username (str):       Device SSH username.
    password (str):       Device SSH password.
    commands (list[str]): Ordered list of exec-mode CLI commands to send.


    RETURN VALUE
    ------------
    Tuple of (index, result) where result is one of:
        'SUCCESS'     — All commands sent without error.
        'BAD_COMMAND' — A command's response contained an IOS error pattern (% Invalid input, etc.).
        'AUTH_BAD'    — Device rejected the credentials.
        'OFFLINE'     — SSH connection could not be established.
        'TIMEOUT'     — Read timeout — device was slow or unresponsive.
        'CONFIG_ERROR'— Syntax or unsupported command detected by Netmiko.
        'SSH_ERROR'   — SSH cipher/key exchange/session failure.
        'UNEXPECTED_ERROR' — Any other unhandled exception.
    """

    ip       = row['OOBM IP Address']
    hostname = row['Hostname']
    index    = row.name

    try:
        connection = ConnectHandler(
            device_type = 'cisco_ios',
            host        = ip,
            username    = username,
            password    = password,
        )

        # Error patterns returned by IOS when a command is invalid
        error_patterns = [
            "% Invalid input",
            "% Incomplete command",
            "% Ambiguous command",
            "% Bad IP address or host name%",
            "Unknown command or computer name, or unable to find computer address"
        ]

        for cmd in commands:
            
            #------------------------------------------------------------------
            logger.info(f"Pushing command(s) to '{hostname}' ({ip}) — '{cmd}'")
            #------------------------------------------------------------------
            output = connection.send_command(cmd, expect_string=r'(?:\([\w\-]+\))?[>#]\s*$', read_timeout=30)

            # # — exec prompt
            # > — unprivileged prompt
            # (config)# — global config
            # (config-if)#, (config-router)#, etc. — sub-mode prompts

            if any(pattern in output for pattern in error_patterns):
                connection.disconnect()
                #---------------------------------------------------------------------------------------
                logger.error(f"Bad command on '{hostname}' ({ip}) — '{cmd}' returned: {output.strip()}")
                #---------------------------------------------------------------------------------------
                return index, 'BAD_COMMAND'

        connection.disconnect()

        #------------------------------------------------------------------
        logger.info(f"Commands pushed successfully to '{hostname}' ({ip})")
        #------------------------------------------------------------------

        return index, 'SUCCESS'
    
    except Exception as e:

        if isinstance(e, NetmikoAuthenticationException):
            #--------------------------------------------------------------------------------
            logger.error(f"Authentication failed on '{hostname}' ({ip}) during command push")
            #--------------------------------------------------------------------------------
            return index, 'AUTH_BAD'

        elif isinstance(e, NetmikoTimeoutException):
            #------------------------------------------------------------------------------
            logger.error(f"Could not reach '{hostname}' ({ip}) for command push (timeout)")
            #------------------------------------------------------------------------------
            return index, 'OFFLINE'

        elif isinstance(e, ReadTimeout):
            #----------------------------------------------------------------------------------------
            logger.error(f"Read timeout on '{hostname}' ({ip}) – device may be slow or unresponsive")
            #----------------------------------------------------------------------------------------
            return index, 'TIMEOUT'

        elif isinstance(e, ConfigInvalidException):
            #------------------------------------------------------------------------------------------
            logger.error(f"Configuration error on '{hostname}' ({ip}) — syntax or unsupported command")
            #------------------------------------------------------------------------------------------
            return index, 'CONFIG_ERROR'

        elif isinstance(e, SSHException):
            #------------------------------------------------------------------------------------
            logger.error(f"SSH error on '{hostname}' ({ip}) — cipher/key exchange/session issue")
            #------------------------------------------------------------------------------------
            return index, 'SSH_ERROR'

        else:
            # Catch-all for any other exception
            #-----------------------------------------------------------------------------------------------------
            logger.error(f"Unexpected error on '{hostname}' ({ip}) during command push — {type(e).__name__}: {e}")
            #-----------------------------------------------------------------------------------------------------
            return index, 'UNEXPECTED_ERROR'
        
# Thread pool -- Called by '' from 'program_gui'
def push_commands_all(selected_df, username, password, commands):

    """
    PURPOSE
    -------
    Concurrently pushes a list of exec-mode CLI commands to all devices in the
    supplied DataFrame using a thread pool. Each device is handled independently
    by push_commands.


    ARGUMENTS
    ---------
    selected_df (pd.DataFrame): DataFrame of the devices selected by the user.
                                Must contain 'OOBM IP Address' and 'Hostname'.
    username    (str):          Device SSH username.
    password    (str):          Device SSH password.
    commands    (list[str]):    Ordered list of exec-mode CLI commands to send.


    RETURN VALUE
    ------------
    List of (index, result) tuples — one per device.
    index matches the original DataFrame index for identification.
    """

    results = []

    with ThreadPoolExecutor(max_workers=10) as executor:

        futures = {
            executor.submit(push_commands, row, username, password, commands): row['Hostname']
            for _, row in selected_df.iterrows()
        }

        for future in as_completed(futures):
            hostname = futures[future]
            try:
                index, result = future.result()
                results.append((index, result))
                #-------------------------------------------------
                logger.info(f"'{hostname}' push result: {result}")
                #-------------------------------------------------
            except Exception as e:
                #----------------------------------------------------------------------
                logger.error(f"Unexpected error pushing commands to '{hostname}': {e}")
                #----------------------------------------------------------------------

    return results
