#requires -version 7

using namespace System.Net.Sockets
using namespace System.Net.Security
using namespace System.Security.Cryptography


function Assert-SslExpiration {
   <#
      .SYNOPSIS
         Gets the number of days until SSL certificate expires.
   #>
   [CmdletBinding()]
   param(
      [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
      [Uri]$Uri
   )

   end {
      try {
         ($cli = [TcpClient]::new()).Connect($Uri, 443)

         $callback = {
            param($sender, $cert, $chain, $errors)
            return $true
         }

         ($ssl = [SslStream]::new($cli.GetStream(), $false, $callback)
         ).AuthenticateAsClient($Uri)

         ($ssl.RemoteCertificate.NotAfter - (Get-Date)).Days
      }
      catch {
         Write-Error $_
      }
      finally {
         ($ssl, $cli).ForEach{if ($_) {$_.Dispose()}}
      }
   }
}
