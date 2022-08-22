#! /bin/sh

mountPassword=`cat ~/mount_password.txt`

mount.ecryptfs /volumeUSB1/usbshare/@Backup1@ /volume1/EncryptedBackup -o key=passphrase:passphrase_passwd=$mountPassword,ecryptfs_cipher=aes,ecryptfs_key_bytes=32,ecryptfs_passthrough=n,no_sig_cache,ecryptfs_enable_filename_crypto=n