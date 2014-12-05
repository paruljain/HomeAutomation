$port = New-Object System.IO.Ports.SerialPort
$sendCommands = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))

$listenerScript = {
    param(
        [System.IO.Ports.SerialPort]$port,
        [System.Management.Automation.Host.PSHost]$parentHost,
        [System.Collections.Queue]$sendCommands
    )

    function ToHex {
        Process {
            $str += ('{0:X2}' -f $_) + ' '
        }
        End { if ($str) { $str.Trim() } }
    }
       
    $rcvBuffer = New-Object System.Collections.ArrayList
    $portLocked = $false

    while ($port.IsOpen) {
        # Is there anything queued to send?
        if ($sendCommands.Count -gt 0 -and !$portLocked) {
            [byte[]]$command = $sendCommands.Dequeue()
            $portLocked = $true
            [void]$port.Write($command, 0, $command.Count)
        }

        if ($port.BytesToRead -gt 0) {
            $bytesToRead = $port.BytesToRead
            $buffer = New-Object byte[] $bytesToRead
            [void]$port.Read($buffer, 0, $bytesToRead)
            $rcvBuffer.AddRange($buffer)
        }
        while ($rcvBuffer.Count -gt 0) {
            switch($rcvBuffer[0]) {
                0x01 { # Zwave frame
                    if ($rcvBuffer.Count -lt 2) { break }
                    if ($rcvBuffer.Count -lt ($rcvBuffer[1] + 2)) { break }
                    
                    # Send ack to controller
                    [void]$port.Write(@(0x06), 0, 1)
                    
                    # Parse the ZWave frame
                    $frame = $rcvBuffer[2..$rcvBuffer[1]]
                    #$parentHost.UI.WriteLine('Got frame: ' + ($frame | ToHex))

                    if ($frame[0] -eq 0 -and $frame[1] -eq 4) { #This is a report
                        $parentHost.Runspace.Events.GenerateEvent('ZWaveReport:' + $frame[3].ToString(), 'ZWaveDriver', $frame, $null)
                    }
                    elseif ($frame[0] -eq 1 -and $frame[1] -eq 0x13) {
                        # Controller has successfully sent the message
                        $portLocked = $false
                    }

                    # Remove the frame from the receive buffer
                    $rcvBuffer.RemoveRange(0, $rcvBuffer[1] + 2)
                }
                default {
                    # Remove single byte Ack, Nack, etc. messages from controller
                    $rcvBuffer.RemoveAt(0)
                }
            }
        }
        Start-Sleep -Milliseconds 10
    }
}

function Connect {
    param(
        [Parameter(Mandatory=$true)][string]$portName
    )
    if ($port.IsOpen) { throw 'Already connected' }
    $port.PortName = $portName
    $port.DataBits = 8
    $port.BaudRate = 115200
    $port.StopBits = 1
    $port.Parity = 'None'
    $port.DtrEnable = $true
    $port.RtsEnable = $true
    $port.Open()

    $newPS = [Powershell]::Create().AddScript($listenerScript).AddParameters(@{port=$port;parentHost=$Host;sendCommands=$sendCommands})
    $job = $newPS.BeginInvoke()
    #$script:Listener = @{ ps = $newPS; job = $job } # Use for debugging
}

function Disconnect {
    $port.Close()
}

function IsConnected {
    $port.IsOpen
}

function ToHex {
    # Converts byte stream to hex string. Used for debugging
    Process {
        $str += ('{0:X2}' -f $_) + ' '
    }
    End { if ($str) { $str.Trim() } }
}

function CheckSum {
    # Computes checksum of a byte stream. Checksum is required to be sent with every ZWave frame
    Begin { [byte]$result = 0xff }
    Process { $result = $result -bxor $_ }
    End { $result }
}

function SendFrame ([byte[]] $payload) {
    # Creates a ZWave frame and queues it for transmit
    if (!$port.IsOpen) { throw 'Comm port is not open' }
    if ($payload.Length -eq 0) { throw 'SendFrame: Nothing to send' }
    if ($payload.Length -gt 254) { throw 'SendFrame: Frame is too long' }

    # Prefix the length of the payload (add 1 for checksum)
    [byte[]]$frame = ,[byte]($payload.Length + 1) + $payload

    # Append the checksum
    $frame += $frame | CheckSum

    # All ZWave frames start with a 0x01
    $frame = ,1 + $frame

    $sendCommands.Enqueue($frame)
}

$FrameType = @{
    'Request' = 0
    'Response' = 1   
}

$Function = @{
    'SendData' = 0x13
}

$Command = @{
    'Set' = 1
    'Get' = 2
}

function GetSensor {
    # Gets the value of a ZWave sensor
    param(
        [Parameter(Mandatory=$true)][byte]$DeviceAddress,
        [Parameter(Mandatory=$true)][byte]$CommandClass,
        [Parameter(Mandatory=$false)][byte]$SensorIndex = 0
    )
    $CmdLength = 3
    SendFrame ($FrameType.Request, $Function.SendData, $DeviceAddress, $CmdLength, $CommandClass, $Command.Get, $SensorIndex)
}

function SetDevice {
    # Sets a ZWave device. Turns a light on or off. Sets the target temperature, etc.
    param(
        [Parameter(Mandatory=$true)][byte]$DeviceAddress,
        [Parameter(Mandatory=$true)][byte]$CommandClass,
        [Parameter(Mandatory=$false)][byte]$SensorIndex = 0,
        [Parameter(Mandatory=$true)][byte[]]$Parameters
    )
    $CmdLength = 3 + $Parameters.Length
    SendFrame (($FrameType.Request, $Function.SendData, $DeviceAddress, $CmdLength, $CommandClass, $Command.Set, $SensorIndex) + $Parameters)
}

Export-ModuleMember -Function Connect, Disconnect, IsConnected, SendFrame, GetSensor, SetDevice
