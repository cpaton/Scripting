$ModulePath = $PSScriptRoot
$DefaultPublicKeyPath = Join-Path -Path $ModulePath -ChildPath "Home-PublicKey.xml"
$DefaultPrivateKeyPath = Join-Path -Path $ModulePath -ChildPath "Home-PrivateKey.xml"
$KeySize = 2048

<#
.SYNOPSIS
Creates a new public/private key pair
#>
function New-Key() {
	param(
		[Parameter( Position = 1, Mandatory = $false )]
		[string] $PrivateKeyFilePath = $( $DefaultPrivateKeyPath ),
		[Parameter( Position = 2, Mandatory = $false )]
		[string] $PublicKeyFilePath = $( $DefaultPublicKeyPath )
	)

	$rsaCryptoService = New-Object System.Security.Cryptography.RSACryptoServiceProvider $KeySize
	$rsaCryptoService.ToXmlString( $false ) | Set-Content -LiteralPath $PublicKeyFilePath -Encoding UTF8
	$rsaCryptoService.ToXmlString( $true ) | Set-Content -LiteralPath $PrivateKeyFilePath -Encoding UTF8
}

<#
.SYNOPSIS
Encrypts a file using the given public key file
#>
function New-EncryptedFile() {
	param(
		[Parameter( Position = 1, Mandatory, ValueFromPipelineByPropertyName )]
		[Alias( "FullName" )]
		[string]
		$PathToEncrypt,
		[Parameter( Position = 2 )]
		[string]
		$PublicKeyPath = $( $DefaultPublicKeyPath ),
		[Parameter( Position = 3 )]
		[string]
		$EncryptedFilePath,
		# Whether to base64 encode the encrypted file contents
		[Parameter()]
		[switch]
		$Base64
	)

	$ErrorActionPreference = "Stop"

	$cleanupActions = @()
	try
	{
		if ( -not ( Test-Path -Path $PathToEncrypt ) ) {
			Write-Error "Cannot find the file that has been requested to be encrypted <$PathToEncrypt>"
			return
		}

		$toEncrypt = Get-Item $PathToEncrypt
		if ($toEncrypt.PSIsContainer)
		{
			Write-Verbose "Input is a folder, creating temporary ZIP archive"
			$folderPath = $PathToEncrypt
			$temporaryZipPath = "$($PathToEncrypt).zip"
			$PathToEncrypt = $temporaryZipPath
			Compress-Archive -Path $folderPath -DestinationPath $PathToEncrypt -Force
			$cleanupActions += { Remove-Item -Path $temporaryZipPath }
		}

		if ( [System.String]::IsNullOrEmpty( $EncryptedFilePath ) ) {
			$EncryptedFilePath = "$PathToEncrypt.encrypted"
		}
		Write-Host "Encrypting to $EncryptedFilePath"

		#
		# Setup the symetric encrption that will be used to encrypt the file contents
		#
		$symetricCryptoService = Get-SymetricCryptoService
		$symetricCryptoService.GenerateKey()
		$symetricCryptoService.GenerateIV()

		#
		# Store the Key and Salt values for the symetric encryption that will be used to encrypt
		# the file into the file encrypted using the public key
		#
		$assymetricCryptoService = Get-AssymetricCryptoService $PublicKeyPath
		$assymetricCryptoService.Encrypt( $symetricCryptoService.Key, $true ) |
			Set-Content -LiteralPath $EncryptedFilePath -AsByteStream
		$assymetricCryptoService.Encrypt( $symetricCryptoService.IV, $true ) |
			Add-Content -LiteralPath $EncryptedFilePath -AsByteStream

		#
		# Encrypt the file contents using the symetric encryption algorithm
		#
		$dataToEncrypt = [System.IO.File]::ReadAllBytes( $PathToEncrypt )
		$numberOfPaddingBytesAddedByEncryptionAlgorithm = ( 8 - ( $dataToEncrypt.Length % 8 ) )
		Write-Verbose "Padding $($numberOfPaddingBytesAddedByEncryptionAlgorithm)"
		$encryptor = $symetricCryptoService.CreateEncryptor()
		$encryptedFileStream = New-Object System.IO.FileStream( $EncryptedFilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Write )
		$encryptedFileStream.Seek( 0, [System.IO.SeekOrigin]::End ) | Out-Null
		$encryptedFileStream.WriteByte( [System.Convert]::ToByte( $numberOfPaddingBytesAddedByEncryptionAlgorithm ) )
		$encryptionStream = New-Object System.Security.Cryptography.CryptoStream( $encryptedFileStream, $encryptor, [System.Security.Cryptography.CryptoStreamMode]::Write )

		try {
			$encryptionStream.Write( $dataToEncrypt, 0, $dataToEncrypt.Length )
			$encryptionStream.Flush()
			$encryptionStream.Close()
		}
		finally {
			$encryptionStream.Dispose()
		}

		if ($Base64)
		{
			Write-Verbose "Converting contents to base64"
			$encryptedBytes = [System.IO.File]::ReadAllBytes( $EncryptedFilePath )
			$base64String = [System.Convert]::ToBase64String( $encryptedBytes )
			Set-Content $EncryptedFilePath -Value $base64String -NoNewline -Encoding ascii
		}

		return Get-ChildItem -Path $EncryptedFilePath
	}
	finally
	{
		$cleanupActions | ForEach-Object { $_.Invoke() }
	}
}

<#
.SYNOPSIS
Decrypts a file using the given private key
#>
function New-DecryptedFile() {
	param(
		[Parameter( Position = 1, Mandatory, ValueFromPipelineByPropertyName )]
		[Alias( "FullName" )]
		[string]
		$PathToDecrypt,
		[Parameter( Position = 2 )]
		[string]
		$PrivateKeyPath = $( $DefaultPrivateKeyPath ),
		[Parameter( Position = 3 )]
		[string]
		$DecryptedFilePath,
		# Specify to denote the encrypted file is encoded in base64
		[Parameter()]
		[switch]
		$Base64
	)

	$ErrorActionPreference = "Stop"

	if ( -not ( Test-Path -Path $PathToDecrypt ) ) {
		Write-Error "Cannot find the file that has been requested to be decrypted <$PathToDecrypt>"
		return
	}

	if ( [System.String]::IsNullOrEmpty( $DecryptedFilePath ) ) {
		$DecryptedFilePath = Join-Path `
			-Path $( Split-Path -Path $PathToDecrypt )`
			-ChildPath $( [System.IO.Path]::GetFileNameWithoutExtension( $PathToDecrypt ) + ".decrypted" )
	}
	Write-Host "Decrypting to $DecryptedFilePath"

	$assymetricCryptoService = Get-AssymetricCryptoService $PrivateKeyPath
	$encryptedDataSize = $assymetricCryptoService.KeySize / 8;

	#
	# Get the data to decrypt
	#
	if ($Base64)
	{
		$base64EncryptedString = Get-Content -Path $PathToDecrypt -Raw -Encoding ascii
		$dataToDecrtyptRawBytes = [System.Convert]::FromBase64String($base64EncryptedString)
	}
	else
	{
		$dataToDecrtyptRawBytes = [System.IO.File]::ReadAllBytes( $PathToDecrypt )
	}

	# From the file contents extract the symetric keys used to encrypt the data
	$encryptedSymetricEncryptionKey = $dataToDecrtyptRawBytes[0..( $encryptedDataSize - 1)]
	$encryptedSymetricEncryptionIV = $dataToDecrtyptRawBytes[$encryptedDataSize..( ( 2 * $encryptedDataSize ) - 1 )]
	$numberOfPaddingBytesAddedByEncryptionAlgorithm = [System.Convert]::ToInt32( $dataToDecrtyptRawBytes[( 2 * $encryptedDataSize )] )
	Write-Verbose "Padding $($numberOfPaddingBytesAddedByEncryptionAlgorithm)"
	$dataToDecrypt = $dataToDecrtyptRawBytes[513..$( $dataToDecrtyptRawBytes.Length - 1 )]

	#
	# Setup the symetric encryption algorithm that will be used to decrypt the data
	#

	$symetricCryptoService = Get-SymetricCryptoService
	$symetricCryptoService.Key = $assymetricCryptoService.Decrypt( $encryptedSymetricEncryptionKey, $true )
	$symetricCryptoService.IV = $assymetricCryptoService.Decrypt( $encryptedSymetricEncryptionIV, $true )

	#
	# Decrypt the file contents using the symetric algorithm
	#
	$decryptor = $symetricCryptoService.CreateDecryptor()
	$encryptedByteStream = New-Object System.IO.MemoryStream( (,$dataToDecrypt) )

	$decryptionStream = New-Object System.Security.Cryptography.CryptoStream( $encryptedByteStream, $decryptor, [System.Security.Cryptography.CryptoStreamMode]::Read )
	$decryptedData = New-Object byte[] $( $dataToDecrypt.Length - $numberOfPaddingBytesAddedByEncryptionAlgorithm )

	$totalBytesRead = 0
	do
	{
		$bytesRead = $decryptionStream.Read( $decryptedData, $totalBytesRead, ( $decryptedData.Length - $totalBytesRead ) )
		$totalBytesRead += $bytesRead
	}
	until ($bytesRead -le 0 -or $totalBytesRead -ge $decryptedData.Length)
	[System.IO.File]::WriteAllBytes( $DecryptedFilePath, $decryptedData )

	return Get-ChildItem -LiteralPath $DecryptedFilePath
}

#
# Private module functions
#
function Get-AssymetricCryptoService( $KeyPath ) {
	$cryptoService = New-Object System.Security.Cryptography.RSACryptoServiceProvider $KeySize
	$keyXml = Get-Content -LiteralPath $KeyPath -Encoding UTF8
	$cryptoService.FromXmlString( $keyXml )

	return $cryptoService
}

function Get-SymetricCryptoService() {
	$cryptoService = New-Object System.Security.Cryptography.TripleDESCryptoServiceProvider
	$cryptoService.KeySize = 192
	return $cryptoService
}

Export-ModuleMember -Function New-Key
Export-ModuleMember -Function New-EncryptedFile
Export-ModuleMember -Function New-DecryptedFile