"""
cml_lab_builder.py
==================
Builds a CML lab with:
  - 40 IOL-XE routers  (TEST-R01 ... TEST-R40)
  - 2 IOL L2-XE switches (TEST-SW01, TEST-SW02)
  - R01-R20 uplinked to TEST-SW01  |  R21-R40 uplinked to TEST-SW02
  - Day-0 config on every router:
      * Hostname
      * enable secret / username admin priv 15 secret
      * SSH v2 (RSA 2048, label-based -- no ip domain-name required)
      * Ethernet0/0: 10.17.1.X/24  (X = 10 + 0-based index, R01=.10 ... R40=.49)
      * ip route 0.0.0.0 0.0.0.0 10.17.1.1

Requires:
    pip install virl2-client

Args (environment variables):
    CML_HOST        HTTPS address of your CML controller
    CML_USER        CML username
    CML_PASS        CML password
    CML_SSL_VERIFY  Set to "true" if using a trusted cert (default: false)

Warning:
    Do NOT hardcode credentials in this file. Use environment variables
    or a .env file excluded from version control.
"""

import os
import sys
import textwrap
import virl2_client

# ── Configuration ─────────────────────────────────────────────────────────────

CML_HOST       = os.environ.get("CML_HOST",     "https://10.211.1.3")
CML_USER       = os.environ.get("CML_USER",     "")   # set via env
CML_PASS       = os.environ.get("CML_PASS",     "")   # set via env
CML_SSL_VERIFY = os.environ.get("CML_SSL_VERIFY", "false").lower() == "true"

LAB_TITLE      = "LAB_NETAUTO002 -- Testing the Tactical-CiscoIOS-Controller"
ROUTER_COUNT   = 40
ROUTER_PREFIX  = "TEST-R"
SWITCH_PREFIX  = "TEST-SW"

ROUTER_DEF     = "iol-xe"
SWITCH_DEF     = "ioll2-xe"

# IP plan
SUBNET         = "10.17.1"
MASK           = "255.255.255.0"
IP_START       = 10       # R01=.10, R02=.11 ... R40=.49
DEFAULT_GW     = "10.17.1.1"
SSH_SECRET     = "admin"

# Canvas layout
GRID_COLS          = 10
ROUTER_X_SPACING   = 150
ROUTER_Y_SPACING   = 120
ROUTER_GRID_START  = (100, 300)
SWITCH_Y           = 100
SWITCH_X_SW01      = 100 + 2 * ROUTER_X_SPACING
SWITCH_X_SW02      = 100 + 7 * ROUTER_X_SPACING


# ── Day-0 config template ─────────────────────────────────────────────────────

def build_day0(hostname: str, ip_address: str) -> str:
    """
    Build the full day-0 startup configuration for one IOL-XE router.

    Args:
        hostname   : Device hostname string (e.g. 'TEST-R01').
        ip_address : IPv4 address for Ethernet0/0 (e.g. '10.17.1.10').

    Returns:
        Multi-line IOS configuration string ready for node.configuration.

    Warning:
        RSA key generation uses the 'label' keyword so no 'ip domain-name'
        is required. Valid on IOS 12.3(14)T+ and IOS-XE.
    """
    return textwrap.dedent(f"""\
        hostname {hostname}
        !
        enable secret {SSH_SECRET}
        !
        username admin privilege 15 secret {SSH_SECRET}
        !
        ip ssh version 2
        !
        crypto key generate rsa label SSH_KEY modulus 2048
        !
        interface Ethernet0/0
         description UPLINK-TO-SWITCH
         ip address {ip_address} {MASK}
         no shutdown
        !
        ip route 0.0.0.0 0.0.0.0 {DEFAULT_GW}
        !
        line vty 0 4
         login local
         transport input ssh
        !
        end
    """)


# ── Helpers ───────────────────────────────────────────────────────────────────

def node_xy(index: int) -> tuple[int, int]:
    """
    Map a 0-based router index to (x, y) canvas coordinates.

    Args:
        index: 0-based router index (0 = TEST-R01).

    Returns:
        (x, y) tuple for the CML canvas.
    """
    col = index % GRID_COLS
    row = index // GRID_COLS
    x = ROUTER_GRID_START[0] + col * ROUTER_X_SPACING
    y = ROUTER_GRID_START[1] + row * ROUTER_Y_SPACING
    return x, y


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if not CML_USER or not CML_PASS:
        sys.exit("[!] CML_USER and CML_PASS must be set as environment variables.")

    print(f"[*] Connecting to CML at {CML_HOST} ...")
    try:
        client = virl2_client.ClientLibrary(
            url=CML_HOST,
            username=CML_USER,
            password=CML_PASS,
            ssl_verify=CML_SSL_VERIFY,
        )
    except Exception as exc:
        sys.exit(f"[!] Connection failed: {exc}")

    try:
        ver = client.system_information().get("version", "unknown")
    except Exception:
        ver = "unknown"
    print(f"[+] Connected  (server version: {ver})")

    # Wipe any existing lab with the same title
    for existing in client.all_labs(show_all=True):
        if existing.title == LAB_TITLE:
            print(f"[!] Removing existing lab '{LAB_TITLE}' (id={existing.id}) ...")
            try:
                existing.stop()
                existing.wipe()
            except Exception:
                pass
            existing.remove()

    # ── Create lab ────────────────────────────────────────────────────────────
    print(f"[*] Creating lab '{LAB_TITLE}' ...")
    lab = client.create_lab(title=LAB_TITLE)
    print(f"[+] Lab created  (id={lab.id})")

    # ── Create switches ───────────────────────────────────────────────────────
    print("[*] Adding switches ...")
    sw01 = lab.create_node(
        label=f"{SWITCH_PREFIX}01",
        node_definition=SWITCH_DEF,
        x=SWITCH_X_SW01,
        y=SWITCH_Y,
    )
    sw02 = lab.create_node(
        label=f"{SWITCH_PREFIX}02",
        node_definition=SWITCH_DEF,
        x=SWITCH_X_SW02,
        y=SWITCH_Y,
    )
    print(f"[+] {sw01.label} and {sw02.label} created")

    # ── Create routers, push day-0 config, and link ───────────────────────────
    print(f"[*] Adding {ROUTER_COUNT} IOL-XE routers ...")
    for i in range(ROUTER_COUNT):
        label      = f"{ROUTER_PREFIX}{i + 1:02d}"
        ip_address = f"{SUBNET}.{IP_START + i}"
        x, y       = node_xy(i)

        router = lab.create_node(
            label=label,
            node_definition=ROUTER_DEF,
            x=x,
            y=y,
        )

        router.configuration = build_day0(label, ip_address)

        target_switch = sw01 if i < 20 else sw02

        # Both sides need a real server-side interface before create_link
        router_iface = router.create_interface()
        switch_iface = target_switch.create_interface()
        lab.create_link(router_iface, switch_iface)

        print(f"  [{i + 1:02d}/{ROUTER_COUNT}]  {label:<12}  {ip_address:<15}  -> {target_switch.label}")

    # ── Summary ───────────────────────────────────────────────────────────────
    lab_url = f"{CML_HOST}/lab/{lab.id}"
    print()
    print("=" * 64)
    print(f"  Lab '{LAB_TITLE}' built successfully!")
    print(f"  Lab ID  : {lab.id}")
    print(f"  URL     : {lab_url}")
    print(f"  Nodes   : {ROUTER_COUNT} routers + 2 switches")
    print(f"  Links   : {ROUTER_COUNT} total")
    print(f"  IPs     : {SUBNET}.{IP_START} - {SUBNET}.{IP_START + ROUTER_COUNT - 1}/24")
    print(f"  Gateway : {DEFAULT_GW}")
    print("=" * 64)


if __name__ == "__main__":
    main()