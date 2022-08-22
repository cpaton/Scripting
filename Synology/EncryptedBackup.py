#! /usr/bin/python

import os
import subprocess
from utils import *
from tempfile import *

init_script(__file__)

def test_writing_file():
    log('Test writing file to {0}'.format(mountPoint))
    with NamedTemporaryFile(dir=mountPoint, delete=True) as testFile:
        encryptedTestFilePath = os.path.join(pathToMount, os.path.basename(testFile.name))
        log('Test file {0} written.  Checking if exists at {1}'.format(testFile.name, encryptedTestFilePath))
        if os.path.exists(encryptedTestFilePath):
            log('{0} exists.'.format(encryptedTestFilePath))
            return True
        else:
            log('Encrypted file {0} was not created.'.format(encryptedTestFilePath))    
            return False

def mount_usb_disk(mountAt, pathToMount):
    passwordFilePath = os.path.expanduser('~/mount_password.txt')

    if not os.path.exists(passwordFilePath):
        raise Exception('Mount password file not found at {0}'.format(passwordFilePath))

    with open(passwordFilePath, 'r') as passwordFile:
        mountPassword = passwordFile.read().replace('\n', '')

    log('Mounting backup device')
    execute_with_log_redirection(['mount.ecryptfs', pathToMount, mountAt, '-o', 'key=passphrase:passphrase_passwd={0},ecryptfs_cipher=aes,ecryptfs_key_bytes=32,ecryptfs_passthrough=n,no_sig_cache,ecryptfs_enable_filename_crypto=n'.format(mountPassword)])

try:

    mountPoint = '/volume1/EncryptedBackup'
    pathToMount = '/volumeUSB1/usbshare/@Backup1@'
    needToMount = False

    foldersToBackup = dict(
        [('/volume1/Data/Backup', 'Backup'),
        ('/volume1/Data/Git', 'Git'),        
        ('/volume1/Data/Movies', 'Movies'),
        ('/volume1/Data/Music', 'Music'),
        ('/volume1/Data/Pictures', 'Pictures'),
        ('/volume1/Data/Repository', 'Repository'),
        ('/volume1/Data/RepositoryCopy', 'RepositoryCopy'),
        ('/volume1/Data/Software', 'Software'),
        ('/volume1/Data/TV', 'TV'),
        ('/volume1/Data/Video', 'Video'),
        ('/volume1/docker', 'docker')])

    pathToMountParent,_ = os.path.split(pathToMount)
    log('Checking {0} is a mounted USB drive'.format(pathToMountParent))
    if not os.path.exists(pathToMountParent):
        raise Exception('USB drive does not appear to be mounted at {0}'.format(pathToMountParent))

    if not os.path.ismount(pathToMountParent):
        raise Exception('USB drive {0} does not appear to be a mount point as expected'.format(pathToMountParent))

    if not os.path.exists(pathToMount):
        raise Exception('Backup target directory {0} does not exist'.format(pathToMount))

    if os.path.exists(mountPoint):
        log('{0} exists'.format(mountPoint))

        # check to see whether it has been mounted
        if not os.path.ismount(mountPoint):
            log('{0} is not a mount point'.format(mountPoint))
            needToMount = True
        else:
            log('{0} is a mount point'.format(mountPoint))

            rawMountOutput = subprocess.check_output(['mount'])
            mountLines = rawMountOutput.split('\n')
            actualMountPoint = None
            mounts = {}
            for line in mountLines:
                mountParts = line.split(' ')
                if len(mountParts) > 2:
                    if mountParts[2] == mountPoint:
                        actualMountPoint = mountParts[0]
            log ('{0} is mounted as {1}'.format(mountPoint, actualMountPoint))
            if actualMountPoint != pathToMount:
                raise Exception('mount point {0} is incorrect'.format(actualMountPoint))
            else:
                log('mount is setup correctly.')
    else:
        log('{0} doesn\'t exist'.format(mountPoint))
        needToMount = True

    if needToMount:
        mount_usb_disk(mountPoint, pathToMount)

    fileWritten = test_writing_file()            
    if not fileWritten:
        execute_script('MountBackup.sh')        
        fileWritten = test_writing_file()
        if not fileWritten:
            raise Exception('Writing test file failed.')        

    log('Starting rsync backup')
    for folderToBackup, relativeBackupFolder in foldersToBackup.iteritems():
        fullBackupPath = os.path.join(mountPoint, relativeBackupFolder)
        log('Backing up {0} to {1}'.format(folderToBackup, fullBackupPath))

        if not os.path.exists(folderToBackup):
            raise Exception('Folder {0} does not exist to be backed up'.format(folderToBackup))

        if not os.path.exists(fullBackupPath):
            log('Creating {0}'.format(fullBackupPath))
            os.mkdir(fullBackupPath)

        execute_with_log_redirection(['rsync', '--recursive', '--delete', '--delete-during', '--links', '--times', '--perms', '--group', '--owner', '--devices', '--verbose', '--stats', '--human-readable', '{0}/'.format(os.path.abspath(folderToBackup)), fullBackupPath])
        
    log('Rsync backup complete')

    execute_with_log_redirection(['sync'])
    execute_with_log_redirection(['sync'])
    execute_with_log_redirection(['sync'])

    log_completed()
except Exception as err:
    notify_error(err)
    raise
