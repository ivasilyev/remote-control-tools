
import logging
import subprocess as sp
from re import findall
from time import sleep
from threading import Thread
from constants import WIN_CMDKEY, WIN_MSTSC, WIN_QWINSTA


def start_rdp_session(server: str, user: str, password: str):
    logging.info(f"Start RDP session to '{user}@{server}'")
    full_server_name = f"TERMSRV/{server}"
    full_user_name = f"{server}\\{user}"

    logging.debug(f"Delete user '{user}' from '{server}'")
    sp.run(f"{WIN_CMDKEY} /delete:{full_server_name}")

    logging.debug(f"Add user '{user}' to '{server}'")
    sp.run(f"{WIN_CMDKEY} /generic:{full_server_name} /user:{full_user_name} /pass:{password}")

    logging.debug(f"Create RDP connection to '{server}'")
    sp.run(f"{WIN_MSTSC} /v:{server} /w:640 /h:480 /noConsentPrompt")

    logging.debug(f"Delete user '{user}' from '{server}'")
    sp.run(f"{WIN_CMDKEY} /delete:{full_server_name}")


def get_session_id(server: str, user: str):
    o = sp.getoutput(f"{WIN_QWINSTA} /server:{server} {user}")
    s = findall("[ \t]+([0-9]+)[ \t]+", o)
    logging.debug(f"STDOUT from 'qwinsta': {o}")
    if len(s) == 0:
        logging.warning("Unable to find Session ID")
        return 0
    out = int(s[-1])
    logging.debug(f"The latest RDP Session ID is {out}")
    return out


def remote_execution(server: str, user: str, password: str, session_id: int, psexec: str, exe: str):
    logging.info(f"Launched remote execution for user'{user}' on '{server}'")
    sp.run(f"{psexec} \\\\{server} -u {user} -p {password} -i {session_id} \"{exe}\"")


def run(netbios_name: str, user: str, password: str, psexec: str, exe: str, wait: int = 5):
    logging.debug("Started remote execution, input parameters: '{}'".format(dict(
        netbios_name=netbios_name,
        user=user,
        password=password,
        psexec=psexec,
        exe=exe
    )))
    thread = Thread(
        target=start_rdp_session,
        kwargs=dict(
            server=netbios_name,
            user=user,
            password=password
        )
    )
    thread.start()
    sleep(wait)
    session_id = get_session_id(netbios_name, user)
    if session_id > 0:
        remote_execution(
            server=netbios_name,
            user=user,
            password=password,
            session_id=session_id,
            psexec=psexec,
            exe=exe,
        )
    else:
        logging.warning("Skip remote execution")
    thread.join()
    logging.debug("Completed remote execution, input parameters: '{}'")


def parse_args():
    from argparse import ArgumentParser
    parser = ArgumentParser(
        description="Tool to execute program using RDP and Sysinternals' PsExec on a remote Windows server",
    )
    parser.add_argument("-s", "--server", metavar="<str>", required=True, type=str,
                        help="Server's NetBIOS name (not IP address!)")
    parser.add_argument("-u", "--user", metavar="<str>", required=True, type=str,
                        help="User name to connect to the server")
    parser.add_argument("-p", "--password", metavar="<str>", required=True, type=str,
                        help="Password to connect to the server")
    parser.add_argument("-c", "--psexec", metavar="<str>", required=True, type=str,
                        help="PsExec executable including full path, e.g. 'C:\\Sysinternals\\PsExec.exe' (on local PC)")
    parser.add_argument("-x", "--exe", metavar="<str>", required=True, type=str,
                        help="Remote executable including full path, e.g. 'C:\\scripts\\startup.bat' (on remote PC)")
    parser.add_argument("-w", "--wait", metavar="<int>", default=5, type=int,
                        help="Seconds to wait while RDP session is established")
    parser.add_argument("-l", "--logging", metavar="<int>", default=2, type=int, choices=list(range(0, 6)),
                        help="Logging level from 0 (ALL) to 5 (CRITICAL), inclusive")
    _namespace = parser.parse_args()
    return (
        _namespace.server,
        _namespace.user,
        _namespace.password,
        _namespace.psexec,
        _namespace.exe,
        _namespace.wait,
        _namespace.logging * 10,
    )


if __name__ == '__main__':
    (
        input_server,
        input_user,
        input_password,
        input_psexec,
        input_exe,
        input_wait,
        input_logging,
    ) = parse_args()

    logging.basicConfig(
        level=input_logging,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )

    run(
        netbios_name=input_server,
        user=input_user,
        password=input_password,
        psexec=input_psexec,
        exe=input_exe,
        wait=input_wait
    )
