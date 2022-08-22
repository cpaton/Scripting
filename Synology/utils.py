import os
import datetime
import subprocess

start_time = datetime.datetime.now()
log_directory = '/volume1/Data/Management/Log'
log_file_handle = None
running_script_path = ''
running_script_name = ''

scripts_directory = os.path.dirname(os.path.realpath(__file__))

def init_script(script_path):
    global log_path
    global running_script_path
    global running_script_name
    global log_file_handle

    script_directory, script_file_name = os.path.split(script_path)
    script_name, script_extension = os.path.splitext(script_file_name)

    running_script_path = script_path
    running_script_name = script_name

    log_name = '{0}-{1:%Y%m%d%H%M}.log'.format(script_name, start_time)
    log_path = os.path.join(log_directory, log_name)

    log_file_handle = open (log_path, 'a')

    log('------------------------------------')
    log('Starting {0} ...'.format(script_path))
    log('------------------------------------')

def log(message):
    message_time = datetime.datetime.now()
    full_message = '{0:%Y-%m-%d %H:%M:%S} {1}'.format(message_time, message)
    print(full_message)
    log_file_handle.write('{0}\n'.format(full_message))
    log_file_handle.flush()

def log_completed():
    log('------------------------------------')
    log('Finshed {0}'.format(running_script_path))
    log('------------------------------------')
    log_file_handle.close()

def get_script_path(file_name):
    return os.path.join(scripts_directory, file_name)

def execute_with_log_redirection(command):
    log(' '.join(str(x) for x in command))
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=log_file_handle)
    lines_iterator = iter(process.stdout.readline, b"")
    for line in lines_iterator:
        log(line) # yield line)

    output = process.communicate()[0]
    exitCode = process.returncode

    if exitCode != 0:
        raise Exception('Failed to run command exit code was {0}'.format(exitCode))


def execute_script(file_name):
    script_to_run = get_script_path(file_name)
    log('Running {0}'.format(script_to_run))
    script_ps = subprocess.Popen(script_to_run, stdout=log_file_handle, stderr=log_file_handle)
    script_ps.communicate()
    if script_ps.returncode != 0:
        raise Exception('Received error code {0} from running {1}'.format(script_ps.returncode, script_to_run))
    else:
        log('{0} completed succesfully'.format(script_to_run))

def notify_error(message):
    log(message)

    log('Sending notification to synology console')
    notification_message = '{0} failed.  Check {1} for more information'.format(running_script_name, log_path)
    subprocess.call(['/usr/syno/bin/synodsmnotify', 'admin', running_script_name, notification_message])

    log('Sending email notification of failure')
    script_ps = subprocess.Popen(['/usr/syno/sbin/synosyslogmail', '--mailtype=SEVERITY', '--severity=ERROR', '--content="{0}"'.format(notification_message)], stdout=log_file_handle, stderr=log_file_handle)
    script_ps.communicate()
    if script_ps.returncode != 0:
        log('Could not send mail for error. Received error code {0}'.format(script_ps.returncode))
    else:
        log('Error email sent.')
