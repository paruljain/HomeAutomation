PowerWave
=========

PowerWave is a pure PowerShell framework for home automation via Z-Wave radio protocol

* Uses Aeon Labs Z-stick to transmit and receive Z-Wave frames
* No dependencies other than PowerShell v3 or better
* Simple to use; great way to learn Z-Wave
* Modular architecture allows Z-Wave devices to be modeled and added independently
* Create powerful scripts to respond to sensor events and control lights, appliances, garage doors, thermostats, etc.
* Combine with PowerShell web server Canister and HTML/Javascript to control your home from any device, from anywhere

Example
-------
Get the heat set point of a Z-Wave enabled 2Gig CT100 thermostat.

    $zdriver = Import-Module .\ZWaveDriver.psm1 -AsCustomObject
    
    # Connect to the Aeon Labs Z-stick on the serial port COM5
    $zdriver.Connect('COM5')
    
    # Assume the device address of the CT100 thermostat is 6
    # In the future the driver module will be enhanced to discover all devices on the network
    $ct100Address = 6
    
    Unregister-Event *
    
    # Create callbacks to process messages sent by Z-Wave devices (called "Reports")
    # One callback for each device address
    Register-EngineEvent -SourceIdentifier "ZWaveReport:$ct100Address" -Action {
        $frame = $Event.SourceArgs
        $cmdClass = $frame[5]
        $sensor = $frame[7]
        $data = $frame[8..($frame[6] + 8)]
        if ($cmdClass -eq 0x43 -and $sensor -eq 1) { Write-Host ('Thermostat Heat Set Point is ' + $data[1].ToString()) }
    }
    
    # Request value of the heat set point from CT100 thermostat
    # 0x43 is the command class for thermostat set points
    # 1 is the heat set point (2 is cooling set point, etc.)
    # See http://220.135.186.178/zwave/example/THERMOSTAT%20SETPOINT/index.html
    
    $zdriver.GetSensor($ct100Address, 0x43, 1)
    
    # After all the work is done release the serial port
    $zdriver.Disconnect()
    
