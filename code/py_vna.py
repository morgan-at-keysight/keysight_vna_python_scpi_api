"""
Keysight VNA SCPI API
Author: Morgan Allison
Updated: 05/2025
Windows 10
Python 3.10.x
PyVISA 1.12.x
Matplotlib 3.4.x
Tested on N5245B PNA-X, P5000A USB VNA, and M9837A PXI VNA

Copyright 2018-2025 Keysight Technologies
All rights reserved

IMPORTANT: This Software includes one or more computer programs bearing
a Keysight copyright notice and in source code format (“Source Files”),
such Source Files are subject to the terms and conditions of the
Keysight `Software End-User License Agreement (“EULA”) <https://www.Keysight.com/find/sweula>`_ and these Supplemental Terms.

BY USING THE SOURCE FILES, YOU AGREE TO BE BOUND BY THE TERMS AND CONDITIONS
OF THE EULA INCLUDING THESE SUPPLEMENTAL TERMS. IF YOU DO NOT AGREE TO
THESE TERMS AND CONDITIONS, DO NOT COPY OR DISTRIBUTE THE SOURCE FILES.

1.  Additional Rights and Limitations. With respect to this Source File,
    Keysight grants you a limited, non-exclusive license, without a right
    to sub-license, to copy, modify and distribute the Source Files
    solely for your internal business purposes or to develop and distribute
    a system or product to which you have added value and only if such
    system or product contains or such internal use utilizes at least one
    Keysight instrument. You own any such modifications and Keysight retains
    all right, title and interest in the underlying Software and Source Files.
    All rights not expressly granted are reserved by Keysight.
2.  Distribution Requirements. Any distribution of the Source Files,
    unmodified or modified, to an external party shall be in conjunction
    with distribution of your system or product and shall be pursuant to
    an enforceable agreement that provides similar protections for Keysight
    and its suppliers as those contained in the EULA and these Supplemental Terms.
3.  General. Capitalized terms used in these Supplemental Terms and not
    otherwise defined herein shall have the meanings assigned to them
    in the EULA. To the extent that any of these Supplemental Terms
    conflict with terms in the EULA, these Supplemental Terms control
    solely with respect to the Source Files.
"""

import matplotlib.pyplot as plt
import numpy as np
import pyvisa
from datetime import datetime, timezone, timedelta
import time

class pyvisaVNA:
    def __init__(self, visaAddress, timeoutMs=10000, openTimeoutMs=100):
        """Class for controlling Keysight VNAs.

        Args:
            visaAddress (str): Visa address of the instrument to be controlled, typically found in Keysight Conection Expert when the instrument is connected
            timeoutMs (int): Timeout value in milliseconds for VISA commands. [default is 10000]
            openTimeoutMs (int): Timeout value in milliseconds when connecting to the resource. [default is 10000]

        Attributes:
            inst (base class for the connected resource): A PyVISA object to be used for communication with the VNA
            instID (str): Instrument information: <company name>, <model number>, <serial number>, <firmware revision>
            instOptions (list): List of all of the instrument options currently installed on the VNA
            numPorts (int): The number of test ports including external testset ports on the VNA
            portCatalog (list): The list of internal test port names on the VNA
            numSources (int): The number of internal sources
            sourceCatalog (list): The list of internal source port names
        """

        # Query alll settings from the VNA and store them as class attributes
        self.inst = pyvisa.ResourceManager().open_resource(visaAddress, open_timeout=openTimeoutMs)
        self.inst.timeout = timeoutMs
        self.instID = self.inst.query('*idn?').rstrip()
        self.instOptions = self.inst.query('*opt?')
        self.numPorts = int(self.inst.query('system:capability:hardware:ports:count?'))
        self.portCatalog = self.inst.query('system:capability:hardware:ports:catalog?').rstrip().strip('"').split(',')
        self.numSources = int(self.inst.query('system:capability:hardware:ports:source:count?'))
        self.sourceCatalog = self.inst.query('system:capability:hardware:ports:source:catalog?').rstrip().strip('"').split(',')

    def close(self):
        """Gracefully closes PyVISA instrument connection."""
        
        self.inst.close()
        del self.inst

    # region Helper Functions
    def print_capabilities(self):
        """Prints instrument information to the console."""

        print(f'Instrument ID: {self.instID}')
        print(f'Port Count: {self.numPorts}')
        print(f'Port Catalog: {self.portCatalog}')
        print(f'Source Count: {self.numSources}')
        print(f'Source Catalog: {self.sourceCatalog}')
        print(f'Options: {self.instOptions}')

    def err_check(self):
        """Prints out all errors and clears error queue. Raises an Exception with the info of the errors encountered."""

        err = []

        # Query errors and remove extra characters
        temp = self.inst.query('syst:err?').strip().replace('+', '').replace('-', '')

        # Read all errors until none are left
        while temp != '0,"No error"':
            # Build list of errors
            err.append(temp)
            temp = self.inst.query('syst:err?').strip().replace('+', '').replace('-', '')
        if err:
            raise Exception(err)

    def source_unleveled_check(self):
        """Checks if a VNA source is unleveled and raises an exception if it is."""
        
        # Bit 2 of the questionable hardware integrity status register indicates if a source is unleveled.
        unleveledState = int(self.inst.query('status:questionable:integrity:hardware:condition?')) & (1 << 2)

        if unleveledState:
            raise ValueError('VNA source is unleveled.')

    def wait_for_opc(self, tempTimeout=None):
        """Waits for the previous command to finish executing before moving on in the script.

        Args:
            tempTimeout (int): Temporary timeout in ms to be used if non-standard timeout is desired. [default is None]
        """
        
        if tempTimeout:
            # If a specified timeout value is entered, temporarily stores the original timeout with the variable originalTimeout, then sets the new specified timeout. After querying opc, the timeout value is set back to the original
            originalTimeout = self.inst.timeout
            self.inst.timeout = tempTimeout
            self.inst.query('*opc?')
            self.inst.timeout = originalTimeout
        else:
            # If NO specified timeout value is entered, the original timeout values will be used for querying opc
            self.inst.query('*opc?')

    def preset(self, clearAll=1):
        """Presets the VNA, and optionally clears all traces and windows.

        Args:
            clearAll (int): Determines whether to clears all windows and traces. [0, 1, default is 1]
        """

        # Clear status register
        self.inst.write('*cls')

        if clearAll:
            # Full preset and remove windows and traces
            self.inst.write('system:fpreset')
        else:
            # Normal preset with S11 trace
            self.inst.write('*rst')
        
        self.wait_for_opc()

    def get_meas_number_from_name(self, measName, ch=1):
        """Gets measurement number from measurement name. This is a helper function used in other class methods and is not intended to be used on its own.

        Args:
            measName (string): Name of the measurement from which to get the corresponding measurement number
            ch (int): Channel to which the measurement belongs. [default is 1]

        Returns:
            int: The measuremnt number of the specified measurement name
        """

        # Select trace
        self.inst.write(f'calculate{ch}:parameter:select "{measName}"')
        
        # Get measurement number of the selected trace
        return int(self.inst.query(f'calculate{ch}:parameter:mnumber?'))

    def get_meas_names(self, ch=1, includeParams=0):
        """Gets all measurement names for a given channel.

        Args:
            ch (int): Channel from which measurement names are queried. [default is 1]
            includeParams (int): Determines whether to return only list of names or a dict of measurement names and measurement parameters. [default is 0]

        Returns:
            if includeParams = 0:
                returns (list): All measurement names. ['mName1', 'mName2']
            else:
                returns (dict): All measurement names and parameters. {'mName1': 'mParam1', 'mName2': 'mParam2'}
        """

        raw = self.inst.query(f'calculate{ch}:parameter:catalog:extended?')
        measurements = raw.strip('"\n').split(',')
        
        measNames = measurements[::2]
        measParams = measurements[1::2]
        
        if includeParams:
            measDict = {} 
            for n, p in zip(measNames, measParams):
                measDict[n] = p
            return measDict
        else:
            return measNames
    
    def select_channel(self, ch=1):
        """Selects the first trace in the specified channel and sets it as active.
        
        Arguments:
        ch (int): Channel to be selected.

        Returns:
            activeChannel (int): Number of the active channel.
        """

        names = self.get_meas_names(ch=ch)
        self.inst.write(f'calculate{ch}:parameter:select "{names[0]}"')
        activeChannel = int(self.inst.query(f'SYSTem:ACTive:CHANnel?').rstrip())

        return activeChannel

    # endregion

    # region Common Functions
    def get_trace(self, measName, ch=1):
        """Acquires frequency and measurement data from selected measurement on VNA for plotting.

        Args:
            measName (str): Measurement from which data will be taken.
            ch (int): Channel to which the measurement belongs. [default is 1]

        Returns:
            freq (NumPy ndArray): Measurement frequency (x-axis) values.
            meas (NumPy ndArray): Measurment data (y-axis) values.
        """
        
        if not isinstance(measName, str):
            raise TypeError('measName must be a string.')

        # Select measurement to be transferred.
        self.inst.write(f'calculate{ch}:parameter:select "{measName}"')

        # Format data for transfer.
        self.inst.write('format:border swap')
        self.inst.write('format real,64')  # Data type is double/float64, not int64.

        # Acquire measurement data.
        meas = self.inst.query_binary_values(f'calculate{ch}:data? fdata', datatype='d')
        self.wait_for_opc()

        # Acquire frequency data.
        freq = self.inst.query_binary_values(f'calculate{ch}:x?', datatype='d')
        self.wait_for_opc()

        return freq, meas

    def get_file(self, vnaPath, remotePath):
        """Transfers a file from "vnaPath" on the VNA to "remotePath" on the remote PC.

        Args:
            vnaPath (str): Full absolute path of the file to be transferred from the VNA to the remote PC.
            remotePath (str): Full absolute path of the destination file location on the remote PC.
        """
        
        # Transfer raw bytes from file on VNA hard drive to remote PC
        dataBytes = self.inst.query_binary_values(f'mmemory:transfer? "{vnaPath}"', datatype='s', container=bytes)
        self.err_check()

        # Write file to remote PC file location
        with open(remotePath, mode='wb') as f:
                    f.write(dataBytes)

    def send_file(self, remotePath, vnaPath):
        """Transfers file from "sourcePath" on the remote PC to "vnaPath" on the VNA.

        Args:
            remotePath (str): Full absolute path of the file on the remote PC that will be transferred to the VNA.
            vnaPath (str): Full absolute path of the destination file location on the VNA.
        """
        
        with open(remotePath, mode='rb') as f:
            rawBytes = f.read()

        # self.inst.write('format:border swapped')
        # Transfer raw bytes from file on remote PC to VNA hard drive
        self.inst.write_binary_values(f'mmemory:transfer "{vnaPath}",', rawBytes, datatype='s')
        
        self.err_check()

    def save_screenshot(self, filePath):
        """Saves a screenshot (MUST BE IN .bmp FORMAT) at destPath on the VNA.
        
        Args:
            filePath (str): Full absolute path of the destination file location on the VNA.
        """
        
        # Save screenshot on VNA hard drive
        self.inst.write(f'mmemory:store:sscreen "{filePath}.bmp"')

    def save_csv(self, measName, fileName, ch=1):
        """Saves csv-formatted measurement data at destPath on the VNA.
        
        Args:
            measName (str): Name of the measurement trace to save.
            fileName (str): Full absolute path of the destination file location on the VNA.
            ch (int): Channel of the data to be saved.
        """
        
        # Select specific trace to be saved
        self.inst.write(f'calculate{ch}:parameter:select "{measName}"')
        # Save csv file on VNA hard drive
        self.inst.write(f'mmemory:store:csv:format "{fileName}"')
        self.err_check()

    def single_trigger(self, tempTimeout=None, ch=1):
        """Executes a single sweep and waits for the sweep to complete.

        Args:
            ch (int): Channel on which the acquisition will be performed. [default is 1]
            tempTimeout (int): Temporary timeout in ms to be used if non-standard timeout is desired. [default is None]
        """

        self.inst.write('trigger:source immediate')
        self.inst.write(f'sense{ch}:sweep:mode single')
        self.wait_for_opc(tempTimeout)

    def hold_trigger(self, ch=1):
        """Sets a given channel to hold so it doesn't run.
        
        Args:
            ch (int): Channel for which the trigger is set to hold. [default is 1]
        """

        self.inst.write(f'trigger:source immediate')
        self.inst.write(f'sense{ch}:sweep:mode hold')

    def marker_activate(self, mkrNum, measName, ch=1):
        """Adds marker to given trace in a given channel.

        Args:
            mkrNum (int): Marker number. [0-15]
            measName (str): Name of the measurement to which the marker will be added.
            ch (int): Channel to which the marker will be added. [default is 1]
        """
        
        # Certain SCPI commands use measurement number instead of measurement name
        # You can get one from the other, thus this helper function
        measNum = self.get_meas_number_from_name(measName, ch)

        self.inst.write(f'calculate{ch}:measure{measNum}:marker{mkrNum}:state 1')

    def marker_format(self, mkrNum, measName, format='default', ch=1):
        """Sets format of marker in a given trace in a given channel.

        Args:
            mkrNum (int): Marker number.
            measName (str): Name of the measurement to which the marker belongs.
            format (str): Format to be applied to the marker. [Default is 'default', see validFormats]
            ch (int): Channel to which the marker belongs.
        """
        
        validFormats = ['default', 'mlinear', 'mlogarithmic', 'impedance', 'admittance', 'phase', 'imaginary', 'real', 'polar', 'gdelay', 'linphase', 'logphase', 'kelvin', 'fahrenheit', 'celsius', 'noise']
        if format.lower() not in validFormats:
            raise ValueError("Invalid 'format', must be 'default', 'mlinear', 'mlogarigthmic', 'impedance', 'admittance', 'phase', 'imaginary', 'real', 'polar', 'gdelay', 'linphase', 'logphase', 'kelvin', 'fahrenheit', 'celsius', or 'noise'.")
        
        # Certain SCPI commands use measurement number instead of measurement name
        # You can get one from the other, thus this helper function
        measNum = self.get_meas_number_from_name(measName, ch)

        self.inst.write(f'calculate{ch}:measure{measNum}:marker{mkrNum}:format {format}')

    def marker_set_x(self, mkrNum, measName, mkrX, ch=1):
        """Sets the x value of the marker.

        Args:
            mkrNum (int): Marker number.
            measName (str): Name of the measurement to which the marker belongs.
            mkrX (float): Desired X value at which the marker will be set.
            ch (int): Channel to which the marker belongs.
        """
        
        # Certain SCPI commands use measurement number instead of measurement name
        # You can get one from the other, thus this helper function
        measNum = self.get_meas_number_from_name(measName, ch)

        self.inst.write(f'calculate{ch}:measure{measNum}:marker{mkrNum}:x {mkrX}')


    def marker_get_x(self, mkrNum, measName, ch=1):
        """Gets the x value of the marker.

        Args:
            mkrNum (int): Marker number.
            measName (str): Name of the measurement to which the marker belongs.
            ch (int): Channel to which the marker belongs.

        Returns:
            mkrX (float): Returns the x value of the selected marker
        """
        
        # Certain SCPI commands use measurement number instead of measurement name
        # You can get one from the other, thus this helper function
        measNum = self.get_meas_number_from_name(measName, ch)

        mkrX = float(self.inst.query(f'calculate{ch}:measure{measNum}:marker{mkrNum}:x?'))

        return mkrX

    def marker_get_y(self, mkrNum, measName, ch=1):
        """Gets the y value of the marker.

        Args:
            mkrNum (int): Marker number.
            measName (str): Name of the measurement to which the marker belongs.
            ch (int): Channel to which the marker belongs.

        Returns:
            mkrY (float): Returns the Y value of the selected marker
        """
        
        # Certain SCPI commands use measurement number instead of measurement name
        # You can get one from the other, thus this helper function
        measNum = self.get_meas_number_from_name(measName, ch)

        # Because a VNA has mag and phase info, there is space for two marker values. 
        # In most cases where log mag is measured, the second value is just 0
        # So we just return the first value
        raw = self.inst.query(f'calculate{ch}:measure{measNum}:marker{mkrNum}:y?').strip()
        mkrY = float(raw.split(',')[0])

        return mkrY

    def add_memory_to_all_traces(self):
        """Saves trace data to memory and changes display to Data and Memory FOR ALL TRACES."""

        # Get all the active channels on the VNA
        rawChannels = self.inst.query(f'system:channels:catalog?').strip('"\n')
        channels = [int(w) for w in rawChannels.split(',')]

        # Iterate through channels, get all traces per channel, and save them to memory
        for ch in channels:
            rawTraces = self.inst.query(f'system:measure:catalog? {ch}').strip('"\n')
            traces = [int(w) for w in rawTraces.split(',')]

            # Save traces to memory
            for t in traces:
                self.inst.write(f'calculate:measure{t}:math:memorize')

        # Change display type to Data and Memory for all traces in each window
        rawWindows = self.inst.query(f'display:catalog?').strip('"\n')
        windows = [int(w) for w in rawWindows.split(',')]

        # Iterate through windows
        for w in windows:
            rawWinTraces = self.inst.query(f'display:window{w}:catalog?').strip('"\n')
            traces = [int(t) for t in rawWinTraces.split(',')]
            
            # Turn data and memory traces 
            for t in traces:
                self.inst.query('*opc?')
                self.inst.write(f'display:window{w}:trace{t}:state on')
                self.inst.write(f'display:window{w}:trace{t}:memory:state on')

    def configure_limit_segment(self, measName, segmentNum, xStart=10e6, xStop=1e9, yStart=-20, yStop=-20, limitType='lmin', ch=1):
        """Cofigures a single limit line segment for a trace. If creating a complex set of limit lines, this method must be called more than once for each trace.
        
        Args:
            measName (str): Name of trace to which limit line is applied.
            segmentNum (int): Limit segment number.
            xStart (float): X-axis start value for limit segment.
            xStop (float): X-axis stop value for limit segment.
            yStart (float): Y-axis start value for limit segment.
            yStop (float): Y-axis stop value for limit segment.
            limitType (str): Type of test the limit checks. 'lmin' will pass a trace with values ABOVE yStart and yStop. 'lmax' will pass a trace with values BELOW yStart and yStop. ['lmin', 'lmax']
            ch (int): Channel to which the limit is applied.
        """
        
        validLimitTypes = ['lmin', 'lmax']
        if limitType.lower() not in validLimitTypes:
            raise ValueError("Invalid 'limitType', must be 'lmin' or 'lmax'.")

        measNum = self.get_meas_number_from_name(measName, ch)

        self.inst.write(f'calculate{ch}:measure{measNum}:limit:segment{segmentNum}:type {limitType}')

        self.inst.write(f'calculate{ch}:measure{measNum}:limit:segment{segmentNum}:stimulus:start {xStart}')
        self.inst.write(f'calculate{ch}:measure{measNum}:limit:segment{segmentNum}:stimulus:stop {xStop}')

        self.inst.write(f'calculate{ch}:measure{measNum}:limit:segment{segmentNum}:amplitude:start {yStart}')
        self.inst.write(f'calculate{ch}:measure{measNum}:limit:segment{segmentNum}:amplitude:stop {yStop}')

    def configure_limit_test(self, measName, limitState=1, limitDisplay=1, limitSound=1, ch=1):
        """Configures a limit test.
        
        Args:
            measName (str): Name of trace to which limit line is applied.
            limitState (int): 1=limit test on, 0=limit test off
            limitDisplay (int): 1=limit line displayed, 0=limit line hidden
            limitSound (int): 1=limit test beeps on failure, 0=no sound on limit test failure
            ch (int): Channel to which the limit is applied.
        """
        
        validLimitStates = [1, 0]
        if limitState not in validLimitStates:
            raise ValueError("Invalid 'limitState', must be 1 or 0.")
        validLimitDisplays = [1, 0]
        if limitDisplay not in validLimitDisplays:
            raise ValueError("Invalid 'limitDisplay', must be 1 or 0.")
        validLimitSounds = [1, 0]
        if limitSound not in validLimitSounds:
            raise ValueError("Invalid 'limitSound', must be 1 or 0.")

        measNum = self.get_meas_number_from_name(measName, ch)

        self.inst.write(f'calculate{ch}:measure{measNum}:limit:state {limitState}')
        self.inst.write(f'calculate{ch}:measure{measNum}:limit:display:state {limitDisplay}')
        self.inst.write(f'calculate{ch}:measure{measNum}:limit:sound:state {limitSound}')

    def get_limit_status(self, measName, ch=1):
        """Returns the pass/fail status of a limit test in a given channel.
        
        Args:
            measName (str): Name of trace to which limit line is applied.
            ch (int): Channel for which to test limits.

        Returns:
            limitStatus (int): 0 is returned when Pass, 1 is returned when Fail.
        """

        measNum = self.get_meas_number_from_name(measName, ch=ch)
        limitStatus = int(self.inst.query(f'calculate{ch}:measure{measNum}:limit:fail?'))
        return limitStatus

    def recall_state_file(self, fileName):
        """Recalls a .csa file.
        
        Args:
            fileName (str): Name of .csa file to recall.
        """

        self.inst.write(f'mmemory:load:csarchive "{fileName}"')
        self.wait_for_opc()
    
    def set_frequency_reference(self, isExtReference=1, refFreq=100e6):
        """Configures the VNA for external reference.

        Args:
            isExtReference (int): Specifies if the reference is external or internal. [0, 1, default is 1]
            refFreq (float): Reference frequency. [default is 100e6]
        """
        
        if isExtReference:
            self.inst.write('sense:roscillator:source external')
            self.wait_for_opc()
            self.inst.write(f'sense:roscillator:external:frequency {refFreq}')
            self.wait_for_opc()
        else:
            self.inst.write('sense:roscillator:source internal')
            self.wait_for_opc()
    
    def configure_receiver_gain(self, sourcePort=1, receieverPort=2, gain='auto', gainCoupling=1, ch=1):
        """
        !!! NOT compatible with PNA or PNA-X. !!!
        !!! NOT compatible with Modulation Distortion or Spectrum Analyzer modes. !!!
        Configures the gain settings for a selected VNA source and receiver port. 

        Args:
            sourcePort (int): Reference source port for which receiver gain is specified. [1, 2, default is 1]
            receiverPort (int): Receiver port for which receiver gain is specified. [1, 2, default is 2]
            gain (str): Receiver gain. ['auto', 'low', 'high', default is 'auto']
            gainCoupling (int): Turns gain coupling on or off. [0, 1, default is 1]
            ch (int): Channel for which receiver gain is configured. [default is 1]
        """
        
        validGains = ['auto', 'low', 'high']
        if gain.lower() not in validGains:
            raise ValueError("Invalid 'gain', must be 'auto', 'low', or 'high'.")
        
        validPorts = [1, 2]
        if sourcePort not in validPorts or receieverPort not in validPorts:
            raise ValueError("Invalid 'sourcePort' or 'receiverPort', must be 1 or 2.")
        
        if not gainCoupling:
            self.inst.write(f'sense{ch}:source:receiver:gain:coupling:all:value off')
        else:
            self.inst.write(f'sense{ch}:source:receiver:gain:coupling:all:value on')
            self.inst.write(f'sense{ch}:source{sourcePort}:receiver{receieverPort}:gain:value "{gain}"')
            
    def configure_receiver_path(self, receiver='b2', rfInSwitch='auto', rfAttenuation=18, ifGain='auto', ch=1):
        """
        !!! Only compatible with E5081A ENA-X and M983xA Configurable PXIe VNAs. !!!
        Configures the RF In Switch, RF Attenuator, and IF Gain settings for a selected VNA receiver. 
        
        Args:
            receiver (str): Receiver for which path is specified. ['a1', 'b1', 'a2', 'b2', default is b2]
            rfInSwitch (str): Switch path for receiver. ['auto', 'internal', 'external', 'bypass20ghz', 'bypass44ghz', 'bypassauto', default is 'auto']
            rfAttenuation (int): Attenuation value for receiver. [0-30 in increments of 2, default is 18]
            ifGain (str/int): Receiver gain. ['auto' or 6-26 in increments of 2, default is 'auto']
            gainCoupling (int): Turns gain coupling off or on. [0, 1, defaults is 0]
            ch (int): Channel for which receiver path is configured. [default is 1]
        """
        
        validReceivers = ['a1', 'b1', 'a2', 'b2']
        if receiver.lower() not in validReceivers:
            raise ValueError("Invalid 'receiver', must be 'a1', 'b1', 'a2', or 'b2'.")
        
        validRfInSwitches = ['auto', 'internal', 'external', 'bypass20ghz', 'bypass44ghz', 'bypassauto']
        if receiver.lower() == 'a1' or receiver.lower() == 'a2':
            # 'a1' and 'a2' receivers cannot be set to the bypass options above
            if rfInSwitch.lower() not in validRfInSwitches[:3]:
                raise ValueError("Invalid 'rfInSwitch', must be 'Auto', 'Internal', or 'External'.")
        else:
            # 'b1' and 'b2' receivers can be set to the bypass options above
            if rfInSwitch.lower() not in validRfInSwitches:
                raise ValueError("Invalid 'rfInSwitch', must be 'Auto', 'Internal', 'External', 'Bypass20GHz', 'Bypass44GHz', or 'BypassAuto'.")
        
        measNames = self.get_meas_names(ch=ch)
        self.inst.write(f'CALCulate{ch}:PARameter:SELect "{measNames[0]}"')
        # print(self.inst.query('system:active:mclass?'))

        port = int(receiver[-1])
        print(port)

        # self.inst.write(f'sense{ch}:source:receiver:gain:coupling:all:value off')
        # self.wait_for_opc()
        print(f'sense{ch}:path:conf:element:select "ReceiverBypass{receiver}","{rfInSwitch}"')
        self.inst.write(f'sense{ch}:path:conf:element:select "ReceiverBypass{receiver}","{rfInSwitch}"')
        self.wait_for_opc()
        self.inst.write(f'source{ch}:power{port}:attenuation:receiver:reference {rfAttenuation}')            
        self.wait_for_opc()
        self.inst.write(f'sense{ch}:path:conf:element:select "IFGain{receiver}","{ifGain}"')
        self.wait_for_opc()
        self.err_check()

    def configure_receiver_leveling(self, port=1, levelingReceiver='r1', levelingType='presweep', maxPower=0, minPower=-100, levelingTolerance=0.1, levelingMaxIterations=5, enableLeveling=1, ch=1):
        """Configures everything in the receiver leveling setup dialog for ONE port. Must be repeated for additional ports.

        Args:
            port (int): Port to which receiver leveling is applied. [1, 2, 3, 4, default is 1]
            levelingReceiver (str): Receiver to be used for leveling. [see validReceivers, default is 'a1']
            levelingType (str): Type of leveling to perform, before or during sweep. ['presweep', 'point', default is 'presweep']
            maxPower (float): Maximum power limit in dBm of the source being leveled. [default is x]
            minPower (float): Minimum power limit in dBm of the source being leveled. [default is x]
            levelingTolerance (float): Tolerance in dB to define if the source is leveled. [default is 0.1]
            levelingMaxIterations (int): Maximum iterations to attempt leveling. [default is 5]
            enableLeveling (int): Enables or disables receiver leveling. [0, 1, default is 1]
            ch (int): Channel for which receiver leveling is enabled. [default is 1]
        """
        
        validPorts = [1, 2, 3, 4, 'VXT']
        if port not in validPorts:
            raise ValueError("Invalid 'port', must be 1, 2, 3, or 4 (or 'VXT' if using MOD).")
        
        validReceivers = ['a', 'b', 'c', 'd', 'r1', 'r2', 'r3', 'r4', 'a1', 'a2', 'a3', 'a4', 'b1', 'b2', 'b3', 'b4']
        if levelingReceiver.lower() not in validReceivers:
            raise ValueError("Invalid 'levelingReceiver', must be 'A', 'B', 'C', 'D', 'R1', 'R2', 'R3', 'R4', 'a1', 'a2', 'a3', 'a4', 'b1', 'b2', 'b3' or 'b4'.")
        
        validLevelingTypes = ['presweep', 'point']
        if levelingType.lower() not in validLevelingTypes:
            raise ValueError("Invalid 'levelingType', must be 'presweep' or 'point'.")
        
        if port == 'VXT':
            self.inst.write(f'source{ch}:power:alc:mode:receiver:state {enableLeveling},"{port}"')
            self.inst.write(f'source{ch}:power:alc:mode:receiver:reference "{levelingReceiver}","{port}"')
            # MOD only allows presweep as a leveling type
            # self.inst.write(f'source{ch}:power:alc:mode:receiver:acquisition:mode {levelingType},"{port}"')
            self.inst.write(f'source{ch}:power:alc:mode:receiver:safe:max {maxPower},"{port}"')
            self.inst.write(f'source{ch}:power:alc:mode:receiver:safe:min {minPower},"{port}"')
            self.inst.write(f'source{ch}:power:alc:mode:receiver:tolerance {levelingTolerance},"{port}"')
            self.inst.write(f'source{ch}:power:alc:mode:receiver:iteration:value {levelingMaxIterations},"{port}"')
            self.inst.write(f'source{ch}:power:alc:mode:receiver:iteration:enable 1,"{port}"')
        else:
            self.inst.write(f'source{ch}:power{port}:alc:mode:receiver:state {enableLeveling}')
            self.inst.write(f'source{ch}:power{port}:alc:mode:receiver:reference "{levelingReceiver}"')
            self.inst.write(f'source{ch}:power{port}:alc:mode:receiver:acquisition:mode {levelingType}')
            self.inst.write(f'source{ch}:power{port}:alc:mode:receiver:safe:max {maxPower}')
            self.inst.write(f'source{ch}:power{port}:alc:mode:receiver:safe:min {minPower}')
            self.inst.write(f'source{ch}:power{port}:alc:mode:receiver:tolerance {levelingTolerance}')
            self.inst.write(f'source{ch}:power{port}:alc:mode:receiver:iteration:value {levelingMaxIterations}')
            self.inst.write(f'source{ch}:power{port}:alc:mode:receiver:iteration:enable 1')
    
    def configure_power_offset(self, port=1, powerOffset=0, ch=1):
        """
        !!! THIS METHOD SHOULD BE USED BEFORE CONFIGURING POWER STIMULUS SETTINGS !!!
        Configures a power offset for a given port in a given channel. 
        Positive values correspond to gain in the path, negative values correspond to loss in the path.

        Args:
            port (int): Port number to which power offset will be applied. [1,2] default is 1.
            powerOffset (float): Power offset in dB that will be applied. Default is 0.
            ch (int): Channel for which power offset will be applied. Default is 1.
        """

        self.inst.write(f'SOURce{ch}:POWer{port}:ALC:MODE:RECeiver:OFFSet {powerOffset}')
    # endregion

    # region Calibration
    def list_cal_sets(self):
        """Lists all available cal sets.
        
        Returns:
            (list): List of available calibration sets.
        """
        
        return self.inst.query('cset:catalog?').rstrip().replace('"','').split(',')

    def load_cal_set(self, calSet, useCalSetStimulus=1, ch=1):
        """Loads an existing calset into the specified VNA measurement channel.
        
        Args:
            calSet (str): Name of existing calset.
            useCalSetStimulus (int): 1=sets stimulus of active channel to match cal set stimulus. [0, 1, default is 1]
            ch (int): Channel to which calibration is applied. Default is 1.
        """

        availableCalSets = self.list_cal_sets()
        if calSet not in availableCalSets:
            raise ValueError('Selected cal set does not exist.')
        
        self.inst.write(f'sense{ch}:correction:cset:activate "{calSet}", {useCalSetStimulus}')
        self.wait_for_opc()
        self.err_check()

    def define_smart_cal(self, numPorts=2, connectors=['APC 3.5 female', 'APC 3.5 female'], calKits=['N4691D User 1 ECal MY57450056', 'N4691D User 1 ECal MY57450056'], ch=1):
        """Defines the DUT connectors and cal kit for Smart Cal.
        
        Args:
            numPorts (int): Number of ports for which the calibration is performed. [1, 2, default is 2]
            connectors (list of strings): Names of the connector types for each DUT port. List length should equal numPorts.
            calKits (list of strings): Names of the cal kit to be used for each DUT port. List length should equal numPorts.
            ch (int): Channel for which calibration is specified. [default is 1]
        """

        # Query list of available connector types
        availableConnectors = self.inst.query(f'sense:correction:collect:guided:connector:cat?').rstrip().strip('"').split(', ')
        # 1.85 mm female, 1.85 mm male, 1.00 mm female, 1.00 mm male, APC 2.4 male, APC 2.4 female, 2.92 mm male, 2.92 mm female,
        # APC 3.5 male, APC 3.5 female, 7-16 male, 7-16 female, APC 7, Type N (50) male, Type N (50) female, Type N (75) male, Type N (75) female,
        # Type F (75) male, Type F (75) female, X-band waveguide, P-band waveguide, K-band waveguide, Q-band waveguide, R-band waveguide, U-band waveguide, 
        # V-band waveguide, W-band waveguide, Type A (50) male, Type A (50) female, Type B
        
        # Check user-specified connectors against available connector types.
        # Check user-specified cal kits against available cal kits.
        for c, k in zip(connectors, calKits):
            if c not in availableConnectors:
                raise ValueError(f'"{c}" is not a valid connector type. Choose from this list: {availableConnectors}.')
            
            # Available cal kits depend on the connector type, so cal kit must be checked per port/connector
            availableCalKits = self.inst.query(f'sense:correction:collect:guided:ckit:cat? "{c}"').rstrip().strip('"').split(', ')
            if k not in availableCalKits:
                raise ValueError(f'"{k}" is not a valid cal kit for {c} connector type. Choose from this list: {availableCalKits}.')
                         
        portArray = range(1, numPorts+1)
        for p, c, k  in zip(portArray, connectors, calKits):
            self.inst.write(f'sense{ch}:correction:collect:guided:connector:port{p} "{c}"')
            self.inst.write(f'sense{ch}:correction:collect:guided:ckit:port{p} "{k}"')

        self.inst.write(f'sense:correction:preference:cset:save user')

    def define_cal_all(self, selectedChannels='all', portArray=[1,2], connectors=['APC 3.5 female', 'APC 3.5 female'], calKits=['N4693D User 1 ECal MY62350179', 'N4693D User 1 ECal MY62350179'], includePowerCal=0, powerCalLevel=0, sparamCalLevel=-15, offsetPorts=[1], powerOffset=[0], saCalPoints=101, psVisa='USB0::0x0957::0x2C18::MY53400002::0::INSTR'):
        """Defines parameters for a Cal All calibration. Note, this does not perform the calibration and should be paired with the run_cal() method to perform the calibration.
        
        Arguments:
            selectedChannels (str): 'all' selects all available channels for calibration. User can input a comma-separated list of channels to be calibrated. e.g. '1, 2, 3'
        """

        # -----------------------------------------------Pre-Req------------------------------------------------------------------
        # Query list of available connector types
        availableConnectors = self.inst.query(f'sense:correction:collect:guided:connector:cat?').rstrip().strip('"').split(', ')
        
        # Check user-specified connectors against available connector types.
        # Check user-specified cal kits against available cal kits.
        for c, k in zip(connectors, calKits):
            if c not in availableConnectors:
                raise ValueError(f'"{c}" is not a valid connector type. Choose from this list: {availableConnectors}.')
            
            # Available cal kits depend on the connector type, so cal kit must be checked per port/connector
            availableCalKits = self.inst.query(f'sense:correction:collect:guided:ckit:cat? "{c}"').rstrip().strip('"').split(', ')
            if k not in availableCalKits:
                raise ValueError(f'"{k}" is not a valid cal kit for {c} connector type. Choose from this list: {availableCalKits}.')


        # -----------------------------------------------Actual Function--------------------------------------------------------------
        # Reset all properties associated with Cal All to default values
        self.inst.write(f'syst:cal:all:reset')

        if selectedChannels == 'all':
            # Select the channels to be calibrated
            # query list of all channels
            # "1,2"

            channelString = self.inst.query(f'SYSTem:CHANnels:CATalog?').rstrip().strip('"')#.split(',')
            # print(channelString)
        else:
            channelString = selectedChannels

        # Select channels to be calibrated
        self.inst.write(f'syst:cal:all:sel {channelString}')

        # may need to add calibration span and calibration points
        self.inst.write(f'SYSTem:CAL:ALL:MCLass:PROPerty:VALue "Calibration Points", "{saCalPoints}"')

        # Specify cal all power offset
        for p, o in zip(offsetPorts, powerOffset):
            self.inst.write(f'SYSTem:CALibrate:ALL:PORT{p}:SOURce:POWer:OFFSet {o}')

        # Specify S-param calibration power
        for p in portArray:
            self.inst.write(f'SYSTem:CALibrate:ALL:PORT{p}:SOURce:POWer {sparamCalLevel}')

        if includePowerCal == 0:
            arg = 'false'
        elif includePowerCal == 1:
            arg = 'true'
        else:
            raise ValueError(f'Invalid argument for includePowerCal f{includePowerCal}.')
        # Power Cal
        self.inst.write(f'SYST:CAL:ALL:MCLass:PROP:VAL "Include Power Calibration", "{arg}"')
        if includePowerCal:
            self.inst.write(f'SYST:COMM:PSEN any, "{psVisa}"')
            self.wait_for_opc()
            self.inst.write(f'SYSTem:CALibrate:ALL:PORT1:SOURce:POWer:VALue {powerCalLevel}')
            self.wait_for_opc()

        # # Not included
        # # Disable Extra Power Cals
        self.inst.write('SYST:CAL:ALL:MCLass:PROP:VAL "Enable Extra Power Cals", "false"')
        # # disable independent calibration channels
        # self.inst.write('SYST:CAL:ALL:MCLass:PROP:VAL "Independent Calibration Channels", "false"')

        # Query the channel number used to measure cal standards
        calChannel = self.inst.query('SYST:CAL:ALL:GUID:CHAN?').rstrip().strip('+')

        for port, connector, ecal in zip(portArray, connectors, calKits):
            self.inst.write(f'sens{calChannel}:corr:coll:guid:conn:port{port} "{connector}"')
            self.inst.write(f'sens{calChannel}:corr:coll:guid:ckit:port{port} "{ecal}"')

        return calChannel

    def measure_cal_standard(self, standardNumber, blocking=1, timeoutMs=30000, ch=1):
        """Measures a given calibration standard. Used in the run_cal() method and is not intended to be used by itself.
        
        Args:
            standardNumber (int): Determines the standard number/calibration step to be measured.
            blocking (bool): 1 causes the measurement SCPI command to be blocking, 0 causes it to be non-blocking. [0, 1, default is 1]
            timeoutMs (int): Temporary timeout setting in milliseconds to allow for long standard measurement times.
            ch (int): Channel on which calibration is performed. [default is 1]
        """
        
        # Acquire individual standard measurements
        # if blocking:
        #     self.inst.write(f'sense{ch}:correction:collect:guided:acquire stan{standardNumber}, synchronous')
        # else:
        #     self.inst.write(f'sense{ch}:correction:collect:guided:acquire stan{standardNumber}, asynchronous')
        
        self.inst.write(f'sense{ch}:correction:collect:guided:acquire stan{standardNumber}')
        
        self.wait_for_opc(tempTimeout=timeoutMs)
        self.err_check()

    def run_cal(self, calSetName='my cal', calSetTimestampSuffix=1, blocking=1, timeoutMs=60000, ch=1, promptUser=1):
        """Initiates a guided calibration. The define_cal_all() method should be run prior to this method.
        
        Args:
            calSetName (str): Name of cal set.
            casetTimestampSuffix (bool): 1=appends a timestamp in the format _YYYYMMdd_hh-mm-ss to the cal set name.
            blocking (int): 1=blocks script from executing until each standard measurement in calibration process completes.
            timeoutMs (int): Temporary timeout setting in milliseconds to allow for long standard measurement times. [default is 30000]
            ch (int): Channel on which calibration is performed. [default is 1]
            promptUser (int): 1 prompts user to make/change connections, 0 just runs the cal without stopping
        """
        
        # Initiates a calibration
        self.inst.write(f'sense{ch}:correction:collect:guided:initiate')

        # Gets the number of steps in the calibration
        numSteps = int(self.inst.query(f'sense{ch}:correction:collect:guided:steps?'))

        # Iterates through calibration steps, prints out connection requirements, 
        # and prompts the user to press enter when the connections have been made
        for s in range(1, numSteps+1):
            print(f'Calibration Step {s} of {numSteps}:')
            print(self.inst.query(f'sense{ch}:correction:collect:guided:description? {s}'))
            if promptUser:
                input('Press enter when connections have been made.')
            print('Measuring cal standard')
            
            successFlag = 0
            while successFlag == 0:
                try:
                    self.measure_cal_standard(s, blocking=blocking, timeoutMs=timeoutMs, ch=ch)
                    # if calibration was successful, everything is connected and the standard has been measured and it is safe to set successFlag to 1
                    successFlag = 1
                    print(f'Completed Calibration Step {s} of {numSteps}.\n\n')
                except Exception as e:
                    print(str(e))
                    input(f'Cal standard not connected correctly. Ensure connection is correct and press enter to try again.')
        print('Calibration Complete.')

        # Generate timestamp suffix in PDT
        if calSetTimestampSuffix:
            # PDT: YYYYMMdd_hh-mm-ss
            # Set up a timezone
            tz = timezone(offset=timedelta(hours=-7), name='pdt')
            # Get current time
            rawTimestamp = datetime.now(tz)
            # Convert datetime object to formatted string
            timestamp = rawTimestamp.strftime('%Y%m%d_%H-%M-%S')

            fullCalSetName = f'{calSetName}_{timestamp}'
        else:
            fullCalSetName = calSetName

        # Complete guided cal, turn on correction, and save cal set
        self.inst.write(f'sense{ch}:correction:collect:guided:save:cset "{fullCalSetName}"')

    def deembed_calset(self, baseCalset, finalCalset, portOneS2p, portTwoS2p, enhancedResponse=1, portOnePowerComp=0, portTwoPowerComp=0, enableExtrapolation=0, overWrite=1):
        """De-embeds TWO s2p files from a base calset and creates a new calset.
        
        Args:
            baseCalset (str): Name of existing calset from which s2p files will be de-embedded.
            finalCalset (str): Name of new calset produced by the de-embedding.
            portOneS2p (str): Absolute file path of s2p at port 1.
            portTwoS2p (str): Absolute file path of s2p at port 2.
            enhancedResponse (int): 1 sets S22 of the fixture to 0 to remove errors from transmission measurements. 0 de-embeds the s2p file normally.
        """

        if overWrite:
            self.inst.write(f'CSET:DELete "{finalCalset}"')

        if enhancedResponse:
            # Creates a new channel and applies the cal set
            deembedChannel = 200
            if "_STD" in baseCalset:
                self.new_sparam_trace(win=deembedChannel, ch=deembedChannel)
            elif "_SMC" in baseCalset:
                self.new_smc_trace(win=deembedChannel, ch=deembedChannel)
            elif "_GCA" in baseCalset:
                self.new_gca_trace(win=deembedChannel, ch=deembedChannel)
            elif "_GCX" in baseCalset:
                self.new_gcax_trace(win=deembedChannel, ch=deembedChannel)
            elif "_NFA" in baseCalset:
                self.new_nf_trace(win=deembedChannel, ch=deembedChannel)
            elif "_NFX" in baseCalset:
                self.new_nfx_trace(win=deembedChannel, ch=deembedChannel)
            elif "_MODX" in baseCalset:
                self.new_modx_trace(win=deembedChannel, ch=deembedChannel)
            elif "_MOD" in baseCalset:
                self.new_mod_trace(win=deembedChannel, ch=deembedChannel)
            self.load_cal_set(baseCalset, useCalSetStimulus=1, ch=deembedChannel)
            self.hold_trigger(ch=deembedChannel)
            
            # Applying MODX calsets does not provide the option to use the calset stimulus, so it must be extracted from the cal set and applied manually
            if "MODX" in baseCalset:
                loFreq = float(self.inst.query(f'cset:frequency:converter? "{baseCalset}", lo1, start').rstrip())
                inputStartFreq = float(self.inst.query(f'cset:frequency:converter? "{baseCalset}", input, start').rstrip())
                inputStopFreq = float(self.inst.query(f'cset:frequency:converter? "{baseCalset}", input, stop').rstrip())
                outputStartFreq = float(self.inst.query(f'cset:frequency:converter? "{baseCalset}", output, start').rstrip())
                outputStopFreq = float(self.inst.query(f'cset:frequency:converter? "{baseCalset}", output, stop').rstrip())
                inputCenterFreq = float((inputStopFreq+inputStartFreq)/2)
                inputSpan = float(inputStopFreq-inputStartFreq)
                outputCenterFreq = float((outputStopFreq+outputStartFreq)/2)
                # print(f'input center: {inputCenterFreq/1e9} GHz')
                # print(f'output center: {outputCenterFreq/1e9} GHz')
                # print(f'input span: {inputSpan/1e9} GHz')
                # print(f'lo: {loFreq/1e9} GHz')

                self.configure_mod_sweep(centerFreq=inputCenterFreq, span=inputSpan, ch=deembedChannel)
                if inputStartFreq < outputStartFreq:
                    sideband = 'high'
                else:
                    sideband = 'low'
                self.configure_modx_mixer(loFreq=loFreq, sideband=sideband, ch=deembedChannel)

            elif "MOD" in baseCalset:
                startFreq = float(self.inst.query(f'cset:frequency:swept? "{baseCalset}", start').rstrip())
                stopFreq = float(self.inst.query(f'cset:frequency:swept? "{baseCalset}", stop').rstrip())
                centerFreq = float((stopFreq+startFreq)/2)
                span = float(stopFreq-startFreq)
                # print(f'center: {centerFreq/1e9} GHz')
                # print(f'span: {span/1e9} GHz')

                self.configure_mod_sweep(centerFreq=centerFreq, span=span, ch=deembedChannel)

            # Add file de-embed to port 1 and zero out reflections at DUT side
            self.inst.write(f'CALCulate{deembedChannel}:FSIMulator:DRAFt:CIRCuit1:ADD FILe, 2')
            self.inst.write(f'CALCulate{deembedChannel}:FSIMulator:DRAFt:CIRCuit1:FILE "{portOneS2p}"')
            self.inst.write(f'CALCulate{deembedChannel}:FSIMulator:DRAFt:CIRCuit1:FILE:MODify NREFLect')
            self.inst.write(f'CALCulate{deembedChannel}:FSIMulator:DRAFt:CIRCuit1:FILE:EXTRapolate {enableExtrapolation}')
            self.inst.write(f'CALCulate{deembedChannel}:FSIMulator:DRAFt:CIRCuit1:VNA:PORTs 1')

            # Add file de-embed to port 2 and zero out reflections at DUT side
            self.inst.write(f'CALCulate{deembedChannel}:FSIMulator:DRAFt:CIRCuit2:ADD FILe, 2')
            self.inst.write(f'CALCulate{deembedChannel}:FSIMulator:DRAFt:CIRCuit2:FILE "{portTwoS2p}"')
            self.inst.write(f'CALCulate{deembedChannel}:FSIMulator:DRAFt:CIRCuit2:FILE:MODify NREFLect')
            self.inst.write(f'CALCulate{deembedChannel}:FSIMulator:DRAFt:CIRCuit2:FILE:EXTRapolate {enableExtrapolation}')
            self.inst.write(f'CALCulate{deembedChannel}:FSIMulator:DRAFt:CIRCuit2:VNA:PORTs 2')
            
            self.inst.write(f'CALCulate{deembedChannel}:FSIMulator:APPLy')
            self.inst.write(f'CALCulate{deembedChannel}:FSIMulator:STATe 1')
            
            if portOnePowerComp:
                self.inst.write(f'CALCulate{deembedChannel}:FSIMulator:POWer:PORT1:COMPensate:STATe 1')
            if portTwoPowerComp:
                self.inst.write(f'CALCulate{deembedChannel}:FSIMulator:POWer:PORT2:COMPensate:STATe 1')
            
            self.inst.write(f'SENSe{deembedChannel}:CORRection:CSET:FLATten "{finalCalset}"')

            self.inst.write(f'system:channels:delete {deembedChannel}')
        else:
            self.inst.write(f'cset:fixture:deembed "{baseCalset}","intermediate","{portOneS2p}",1,1,0')
            self.wait_for_opc()
            self.inst.write(f'cset:fixture:deembed "intermediate","{finalCalset}","{portTwoS2p}",2,1,0')
            self.wait_for_opc()
        
        self.err_check()

    # def modify_s2p_file(self, originalFileName, modifiedFileName, zeroS11=0, zeroS22=0):
    #     """Modifies an existing s2p file, setting S11, S22, or both S11 + S22 to zero.
        
    #     Arguments:
    #         originalFileName (str): full absolute path to the s2p file that will be modified
    #         modifiedFileName (str): full absolute path to the location that the modified s2p file will be saved
    #         zeroS11 (int): 0 = does nothing, 1 = sets S11 of the s2p file to 0 (a perfect match)
    #         zeroS22 (int): 0 = does nothing, 1 = sets S22 of the s2p file to 0 (a perfect match)
    #     """
    #     pass

    def deembed_s2p_file(self, vnaPort, s2pFileName, reverseS2p=0, snnZero=0, enableExtrapolation=0, ch=1):
        """De-embeds an s2p file from specified VNA port. If multiple s2p files must be de-embedded, this method must be called once for each de-embed.
        
        Arguments:
            vnaPort (int): Specifies the VNA port from which the s2p file will be de-embedded. [1,2,3,4]
            s2pFileName (str): Full absolute file path to the s2p being de-embedded.
            reverseS2p (int): 0 leaves the s2p file definition as is, 1 reverses the ports of the s2p file.
            snnZero (int): 0 does not modify s2p file, 1 sets the match at the DUT port to 0 (linear).
            enableExtrapolation (int): 0 does not enable extrapolation, 1 enables extrapolation.
            ch (int): Channel to which de-embedding is applied.
        """

        self.inst.write(f'CALCulate{ch}:FSIMulator:DRAFt:CIRCuit:RESet')
        circuit = int(self.inst.query(f'CALCulate{ch}:FSIMulator:DRAFt:CIRCuit:NEXT?').strip())
        self.inst.write(f'CALCulate{ch}:FSIMulator:DRAFt:CIRCuit{circuit}:ADD FILe,2')
        self.inst.write(f'CALCulate{ch}:FSIMulator:DRAFt:CIRCuit{circuit}:VNA:PORTs {vnaPort}')
        self.inst.write(f'CALCulate{ch}:FSIMulator:DRAFt:CIRCuit{circuit}:FILE "{s2pFileName}"')
        if snnZero:
            self.inst.write(f'CALCulate{ch}:FSIMulator:DRAFt:CIRCuit{circuit}:FILE:MODify NREFlect')
        if reverseS2p:
            self.inst.write(f'CALCulate{ch}:FSIMulator:DRAFt:CIRCuit{circuit}:DEVice:PORTs:REVerse 1')
        if enableExtrapolation:
            self.inst.write(f'CALCulate{ch}:FSIMulator:DRAFt:CIRCuit{circuit}:FILE:EXTRapolate 1')

        self.inst.write(f'CALCulate{ch}:FSIMulator:DRAFt:CIRCuit{circuit}:STATe 1')
        self.inst.write(f'CALCulate{ch}:FSIMulator:DRAFt:APPLy')
        self.inst.write(f'CALCulate{ch}:FSIMulator:STATe 1')

        self.err_check()

    # endregion

    # region ECal
    
    """
     _    _                  _             
    | |  | |                (_)            
    | |  | | __ _ _ __ _ __  _ _ __   __ _ 
    | |/\| |/ _` | '__| '_ \| | '_ \ / _` |
    \  /\  / (_| | |  | | | | | | | | (_| |
     \/  \/ \__,_|_|  |_| |_|_|_| |_|\__, |
                                      __/ |
                                     |___/ 
    """

    """!!! WARNING !!! This entire section is experimental. !!! WARNING !!!"""


    def get_ecal_module_nums(self):
        """Returns a list of ECal module numbers.
        
        Returns:
            (list): List of ints containing ECal module numbers.
        """

        # Grab list of ECal module numbers
        raw = self.inst.query('sense:correction:ckit:ecal:list?').strip().replace('+','').split(',')
        # Convert strings to integers
        modList = [int(r) for r in raw]

        if modList == [0]:
            raise AttributeError('No ECal modules connected.')
        else:
            return modList

    def get_ecal_module_states(self, ecalNum=1, numPorts=2):
        """Returns a list of paths in selected ECal.
        
        Args:
            ecalNum (str): Specifies target ECal module number. [default is 1]
            numPorts (int): Number of physical ports on ECal. [default is 2]
        """

        # CONTrol:ECAL:MODule<num>:PATH:COUNt? <name> tells how many paths the ecal has
        if numPorts == 2:
            portPaths = ['a', 'b', 'ab']
        elif numPorts == 4:
            portPaths = ['a', 'b', 'ab', 'c', 'd', 'ac', 'ad', 'bc', 'bd', 'cd']
        else:
            raise ValueError("Invalid 'numPorts', must be 2 or 4")
        
        # Create dict for storing ECal path info
        stateDict = {}

        # Iterate through paths and get number of states for each path
        for p in portPaths:
            numStates = int(self.inst.query(f'sense:correction:ckit:ecal{ecalNum}:path:count? {p}').strip().replace('+',''))
            stateDict[f'{p}'] = numStates
        
        return stateDict

    def get_individual_ecal_info(self, ecalNum=1):
        """Gets a dict containing info about a given ECal module.
        
        Args:
            ecalNum (str): Specifies target ECal module number. [default is 1]
        """
        
        # Create a dict to collect info about ECal module
        ecalInfoDict = {}

        # Get list of ECal characterizations sests (Char0 is the factory characterization, any others are user characterizations)
        charNum = self.inst.query(f'sense:correction:ckit:ecal{ecalNum}:clist?').strip().replace('+','').split(',')
        
        # Iterate through characterizations for a given ECal module
        for c in charNum:
            raw = self.inst.query(f'sense:correction:ckit:ecal{ecalNum}:information? char{c}').strip().replace('"','').split(', ')
            
            # Create temporary dict to collect info about a specific characterization of a specific ECal module
            rawDict = {}

            # Iterate through all the info and separate the string into keys and values to populate the ECal info
            for r in raw:
                key, value = r.split(': ')
                rawDict[key] = value
            
            # Populate individualEcalDict with information
            ecalInfoDict[f'Char{c}'] = rawDict
        
        return ecalInfoDict

    def get_individual_ecal_model_serial(self, ecalNum=1):
        """Gets a specified ECal model serial number.

        Args:
            ecalNum (str): Specifies target ECal module number. [default is 1]

        Returns:
            string: Serial number of the E-Cal. 
        """
        infoDict = self.get_individual_ecal_info(ecalNum)
        
        return f"{infoDict['Char0']['ModelNumber']},{infoDict['Char0']['SerialNumber']}"

    def get_all_ecal_model_serial(self):
        """Gets the model and serial number for ALL connected ECals.

        Returns:
            list: returns a list of serial numbers of each E-Cal
        """

        modelSerial = []

        for e in self.get_ecal_module_nums():
            modelSerial.append(self.get_individual_ecal_model_serial(e))
        
        return modelSerial

    def get_all_ecal_info(self):
        """Gets full info about connected ECals as a list of dicts."""

        # Create list that will contain information on all ECal modules
        ecalInfo = []

        # Get list of ECal modules
        ecalNums = self.get_ecal_module_nums()
        
        # Iterate through the connected ECal modules
        for e in ecalNums:
            individualEcalDict = self.get_individual_ecal_info(e)
            
            # Populate outer list with information for each individual ECal module
            ecalInfo.append(individualEcalDict)

        return ecalInfo
    
    def get_ecal_sparam_data(self, ecalNum=1, char=0, debug=0):
        """
        ################VERY EXPERIMENTAL################
        
        Gets s-parameter data from ECal.

        Args:
            ecalNum (int, optional): Specifies the E-Cal number that want to get S-Parameter data from. Defaults to 1.
            char (int, optional): _description_. Defaults to 0.
            debug (bool, optional): _description_. Defaults to False.

        Raises:
            ValueError: Raise an error when entered ecalNum is not connected

        Returns:
            list: frequency and their corresponding S-Parameter values
        """

    


        """CHANGE THIS SO THAT IT'S JUST RETURNING S PARAMS FOR ONE PATH AT A TIME
        MANAGING ALL THE DIFFERENT OPTIONS FOR PATHS AND STATES IS WAY TOO MUCH"""





        # SENSe<ch>:CORRection:CKIT:ECAL<mod>:PATH:DATA? <path>, <stateNum>[,<char>] https://na.support.keysight.com/vna/help/latest/Programming/GP-IB_Command_Finder/Sense/CorrCKIT_SCPI.htm#EcalPathData
        # data is returned in the same format as CALCulate<cnum>:MEASure<mnum>:DATA:SNP:PORTs? <"x,y,z".>
        # find ecal information with SENSe:CORRection:CKIT:ECAL<mod>:LIST? and SENSe:CORRection:CKIT:ECAL<mod>:INFormation? [<char>]
        # CONTrol:ECAL:MODule<num>:PATH:COUNt? <name> tells how many paths the ecal has
        
        # Error checking
        if ecalNum not in self.get_ecal_module_nums():
            raise ValueError('"ecalNum" referencing an ECal outside of range of connected ECals.')

        # Get info about ECal
        pathDict = self.get_ecal_module_states(ecalNum)
        infoDict = self.get_individual_ecal_info(ecalNum)
        
        # Set formatting for s-param data
        self.inst.write('format:border swap')
        self.inst.write('format real,64')
        # self.inst.write('format ascii,0')

        numPoints = int(infoDict[f'Char{char}']['NumberOfPoints'])
        print(numPoints)
        # print(len(raw))
        # numPoints = int(len(raw) / 3)
                
        for path, stateValue in pathDict.items():
            for state in range(stateValue):
                raw = np.array(self.inst.query_binary_values(f'sense:correction:ckit:ecal{ecalNum}:path:data? {path},{state},char{char}', datatype='d'))
                freq = raw[:numPoints]
                s11Real = raw[numPoints:int(2 * numPoints)]
                s11Imag = raw[int(2 * numPoints):int(3 * numPoints)]
                s11 = s11Real + 1j * s11Imag
                if path == 'ab':
                    s12Real = raw[int(3 * numPoints):int(4 * numPoints)]
                    s12Imag = raw[int(4 * numPoints):int(5 * numPoints)]
                    s12 = s12Real + 1j * s12Imag
                    s21Real = raw[int(5 * numPoints):int(6 * numPoints)]
                    s21Imag = raw[int(6 * numPoints):int(7 * numPoints)]
                    s21 = s21Real + 1j * s21Imag
                    s22Real = raw[int(7 * numPoints):int(8 * numPoints)]
                    s22Imag = raw[int(8 * numPoints):int(9 * numPoints)]
                    s22 = s22Real + 1j * s22Imag

                if debug:
                    # print(infoDict[f'Char{char}'])
                    # print(type(raw))
                    self.set_ecal_path(ecalNum, path=path, state=state)
                    import matplotlib.pyplot as plt
                    plt.plot(freq, 20 * np.log10(abs(s11)))
                    if path == 'ab':
                        plt.plot(freq, 20 * np.log10(abs(s12)))
                        plt.plot(freq, 20 * np.log10(abs(s21)))
                        plt.plot(freq, 20 * np.log10(abs(s22)))
                    plt.show()
        if path == 'ab':
            pass
        return freq, s11

    def set_ecal_path(self, ecalNum=1, path='a', state=1, numPorts=2):
        """Sets the state of a given path in an ECal.

        Args:
            ecalNum (str): Specifies target ECal module number. [default is 1]
            path (str): Path through the ECal. [default is 'a']
            state (int): Impedance state for the specified path through the ECal. [default is 1]
            numPorts (int): Number of ports in the ECal. [default is 2]
        """

        if ecalNum not in self.get_ecal_module_nums():
            raise ValueError('"ecalNum" referencing an ECal outside of range of connected ECals.')
        
        # if numPorts == 2:
        #     portPaths = ['a', 'b', 'ab']
        # elif numPorts == 4:
        #     portPaths = ['a', 'b', 'ab', 'c', 'd', 'ac', 'ad', 'bc', 'bd', 'cd']
        # else:
        #     raise ValueError("Invalid 'numPorts', must be 2 or 4")


        self.inst.write(f'control:ecal:module{ecalNum}:path:state {path}, {state}')


        # Set the ECal path/state with CONTrol:ECAL:MODule<num>:PATH:STATe <path>, <stateNum> https://na.support.keysight.com/pna/help/latest/Programming/GP-IB_Command_Finder/Control.htm#EcalPathState
        self.inst.write(f'control:ecal:module{ecalNum}:path:state {path}, {state}')
    
    # endregion

    # region Standard Channel
    def new_sparam_trace(self, measName='ReturnLoss', measParam='S11', win=1, ch=1):
        """Creates a new S-parameter trace in a standard channel.
        
        Args:
            measName (str): Name of trace/measurement, e.g. "MyVeryCoolS21Measurement". [default is 'ReturnLoss']
            measParam (str): Name of parameter to be measured by trace, e.g. "S21". [default is 'S11']
            win (int): Window where the trace/measurement will be displayed. [default is 1]
            ch (int): Channel to which the trace/measurement will be assigned. [default is 1]
        """
        
        # validParams = ['S11', 'S12', 'S21', 'S22', 'A, 1', 'A, 2', 'B, 1', 'B, 2', 'R1, 1', 'R2, 2']
        # if measParam not in validParams:
            # raise ValueError("Invalid 'measParam', check measurement parameter argument.")

        # Create new trace with name, parameter, and channel
        self.inst.write(f'calc{ch}:parameter:define:extended "{measName}", "{measParam}"')
        
        # Check for windows
        windows = self.inst.query('display:catalog?').split(',')
        
        # The VNA returns *everything* as a string, so to compare the desired window to the available windows in the VNA, we have to convert the 'win' argument to a string
        # 1 in ['1', '2'] would return False, but '1' in ['1', '2'] will return True, which is what we want.
        
        # If the specified window doesn't exist, create it by turning it on
        if str(win) not in windows:
            self.inst.write(f'display:window{win}:state on')

        # See how many traces are in the window
        traces = self.inst.query(f"display:window{win}:catalog?").strip()

        # Determine the number of the 'next' trace in the window
        # VNA returns the string 'EMPTY' if a window has no traces defined in it
        if 'empty' in traces.lower():
            nextTrace = 1
        else:
            nextTrace = len(traces.split(',')) + 1

        # Feed the new trace into a window and check for errors
        self.inst.write(f'display:window{win}:trace{nextTrace}:feed "{measName}"')
        
        self.wait_for_opc()
        self.err_check()

    def configure_sparam_stimulus(self, startFreq=10e6, stopFreq=44e9, portPower=-15, numPoints=201, ifBw=100e3, ch=1):
        """Configure the frequency, power, number of points, and IF bandwidth for an S-param measurement and pause measurement.
        
        Args:
            startFreq (float): Start frequency in Hz. [default is x]
            stopFreq (float): Stop frequency in Hz. [default is x]
            portPower (float): VNA output power in dBm. [default is x]
            numPoints (int): Number of points to measure. [default is 201]
            ifBw (int): IF Bandwidth in Hz. [default is 100e3]
            ch (int): Channel for which stimulus is configured. [default is 1]
        """

        self.inst.write(f'sense{ch}:frequency:start {startFreq}')
        self.inst.write(f'sense{ch}:frequency:stop {stopFreq}')
        self.inst.write(f'source{ch}:power {portPower}')
        self.inst.write(f'sense{ch}:sweep:points {numPoints}')
        self.inst.write(f'sense{ch}:bandwidth {ifBw}')

        self.inst.write('initiate:continuous 0')

        self.wait_for_opc()

    def save_s2p(self, measName, fileName, ports=[1,2], ch=1):
        """Saves an s2p file from a standard channel measurement.

        Args:
            measName (str): Name of trace/measurement, e.g. "MyVeryCoolS21Measurement".
            fileName (str): Absolute file path for saved s2p file.
            ports (list): List of 2 VNA ports from which data will be saved. [default is [1,2]]
            ch (int): Channel from which data will be saved. [default is 1]
        """
        
        # Certain SCPI commands use measurement number instead of measurement name
        # You can get one from the other, thus this helper function
        measNum = self.get_meas_number_from_name(measName, ch)

        portString = ','.join(map(str, ports))

        self.inst.write(f'calculate{ch}:measure{measNum}:data:snp:ports:save "{portString}","{fileName}"')

    def load_s2p(self, fileName):
        """Loads s2p data into a new channel in the VNA from an external file.
        
        Args:
            fileName (str): Full path of s2p file to load.
        """
        
        self.inst.write(f'mmemory:load "{fileName}"')
        self.wait_for_opc()

    # endregion

    # region Modulation Distortion
    def new_mod_trace(self, measName='InputPower', measParam='PIn1', win=1, ch=1):
        """Creates a new trace in a Modulation Distortion channel.
        
        Args:
            measName (str): Name of the trace to be added. [default is 'InputPower']
            measParam (str): Specifies the trace to be added. [see validParams, default is 'PIn1']
            win (int): Window to which trace will be added. [default is 1]
            ch (int): Channel to which trace will be added. [default is 1]
        """
        
        validParams = ['PIn1', 'POut2', 'PModFile', 'MSig2', 'MDist2', 'MDistIR2', 'MGain21', 
                        'PGain21', 'MComp21', 'PGain21', 'LMatch2', 'CarrIn1', 'CarrOut2', 'NPRIn1', 'NPROut2', 
                        'NPRDist21', 'NPRPwrOut2', 'ACPIn1', 'ACPOut2', 'ACPDist21', 'ACPPwrIn1', 'ACPPwrOut2', 
                        'EVMDistEq21', 'EVMDistUn21', 'EVMPwrIn1', 'EVMPwrOut2', 'ModFilter', 
                        'A', 'b1', 'B', 'C', 'b3', 'D', 'b4', 'R1', 'a1', 'R2', 'a2', 'R3', 'a3', 'R4', 'a4',
                        'S11', 'S21', 'LPIn1', 'LPOut1', 'LPOut2',
                        'PIn2', 'POut1', 'MSig1', 'MDist1', 'MDistIR1', 'MGain12', 
                        'PGain12', 'MComp12', 'PGain12', 'LMatch1', 'CarrIn2', 'CarrOut1', 'NPRIn2', 'NPROut1', 
                        'NPRDist12', 'NPRPwrOut1', 'ACPIn2', 'ACPOut1', 'ACPDist12', 'ACPPwrIn2', 'ACPPwrOut1', 
                        'EVMDistEq12', 'EVMDistUn12', 'EVMPwrIn2', 'EVMPwrOut1']
        
        if measParam not in validParams:
            raise ValueError(f"Invalid 'measParam': {measParam}, check measurement parameter argument.")

        # Create new trace with name, parameter, and channel
        self.inst.write(f'calc{ch}:custom:define "{measName}", "Modulation Distortion", "{measParam}"')
        
        # Check for windows
        windows = self.inst.query('display:catalog?').split(',')
        
        # The VNA returns *everything* as a string, so to compare the desired window to the available windows in the VNA, we have to convert the 'win' argument to a string
        # 1 in ['1', '2'] would return False, but '1' in ['1', '2'] will return True, which is what we want.
        
        # If the specified window doesn't exist, create it by turning it on
        if str(win) not in windows:
            self.inst.write(f'display:window{win}:state on')

        # See how many traces are in the window
        traces = self.inst.query(f"display:window{win}:catalog?").strip()

        # Determine the number of the 'next' trace in the window
        # VNA returns the string 'EMPTY' if a window has no traces defined in it
        if 'empty' in traces.lower():
            nextTrace = 1
        else:
            nextTrace = len(traces.split(',')) + 1

        # Feed the new trace into a window and check for errors
        self.inst.write(f'display:window{win}:trace{nextTrace}:feed "{measName}"')
        
        self.wait_for_opc()
        self.err_check()

    def configure_mod_sweep(self, sweepType='fixed', centerFreq=1e9, span=300e6, noiseBw=200, power=-30, powerContext='din1', startPower=-20, stopPower=-10, powerPoints=11, autoIncreaseNoiseBw=0, ch=1):
        """Configures the Sweep tab in Modulation Distortion Setup.
        
        Args:
            sweepType (str): Type of sweep, fixed frequency or power. ['fixed', 'power', default is 'fixed]
            centerFreq (float): Center frequency in Hz for the modulation distortion measurement. [default is x]
            span (float): Span in Hz for the modulation distortion measurement. [default is x]
            power (float): Power in dBm for the modulation distortion measurement. [default is x]
            powerContext (str): Specifies at which DUT port the specified power should be applied. ['din1', 'dout2', default is 'din1']
            startPower (float): Start power for power sweep types. Ignored for 'fixed' sweepType. [default is x]
            stopPower (float): Stop power for power sweep types. Ignored for 'fixed' sweepType. [default is x]
            powerPoints (int): Specifies how many points to include in the power sweep. Ignored for 'fixed' sweepType. [default is 11]
            autoIncreaseNoiseBw (int): Specifies whether to automatically increase noise bandwidth as power increases in power sweeps. Ignored for 'fixed' sweepType. [0, 1, default is 0]
            ch (int): Channel for which settings are configured. [default is 1]
        """

        validSweepTypes = ['fixed', 'power']
        if sweepType.lower() not in validSweepTypes:
            raise ValueError("Invalid 'sweepType', must be 'fixed' or 'power'.")
        
        self.inst.write(f'sense{ch}:distortion:sweep:type {sweepType}')
        if sweepType.lower() == 'power':
            self.inst.write(f'sense{ch}:distortion:sweep:power:carrier:ramp:level:start {startPower}')
            self.inst.write(f'sense{ch}:distortion:sweep:power:carrier:ramp:level:stop {stopPower}')
            self.inst.write(f'sense{ch}:distortion:sweep:power:carrier:ramp:points {powerPoints}')
            self.inst.write(f'sense{ch}:distortion:sweep:power:carrier:list:nbw {noiseBw}')
            self.inst.write(f'sense{ch}:distortion:sweep:power:carrier:ramp:nbw:auto {autoIncreaseNoiseBw}')

        self.inst.write(f'sense{ch}:distortion:sweep:carrier:frequency {centerFreq}')
        self.inst.write(f'sense{ch}:frequency:span {span}')
        self.inst.write(f'sense{ch}:sa:bandwidth:noise {noiseBw}')
        
        validPowerContexts = ['din1', 'dout2']
        if powerContext.lower() not in validPowerContexts:
            raise ValueError("Invalid 'powerContext', must be 'din1' or 'dout2'.")

        self.inst.write(f'sense{ch}:distortion:sweep:power:carrier:level:port {powerContext}')
        self.inst.write(f'sense{ch}:distortion:sweep:power:carrier:level {power}')
        self.wait_for_opc()

    def configure_mod_rfpath(self, srcAtten=0, includeSrcAtten=1, inputPort=1, outputPort=2, recvAtten=0, ch=1):
        """Configures the RF Path tab in Modulation Distortion.
        
        Args:
            srcAtten (int): Source attenuation in dB. [0, 10, 20, 30, 40, default is 0]
            includeSrcAtten (int): Specifies whether to include source attenuation in power calculations. [0, 1, default is 1]
            inputPort (int): Specifies which VNA port corresponds to DUT input port. [1, 2, default is 1]
            outputPort (int): Specifies which VNA port corresponds to DUT output port. [1, 2, default is 2]
            recvAtten (int): Receiver attenuation in dB. [0, 10, 20, 30, 40, default is 0]
        """
        
        validPorts = [1, 2]
        if inputPort not in validPorts or outputPort not in validPorts:
            raise ValueError("Invalid 'inputPort' or 'outputPort', must be 1 or 2.")
        
        # There's an issue with the source attenuation command if the source port is not 1
        self.inst.write(f'source{ch}:power{inputPort}:attenuation {srcAtten}')

        self.inst.write(f'sense{ch}:distortion:path:source:attenuation:include {includeSrcAtten}')
        self.inst.write(f'sense{ch}:distortion:path:dut:input {inputPort}')
        self.inst.write(f'sense{ch}:distortion:path:dut:output {outputPort}')
        self.inst.write(f'source{ch}:power{outputPort}:attenuation:receiver:test {recvAtten}')
        self.wait_for_opc()

        self.err_check()

    def add_mod_source(self, sourceName='VXT', driver='VXT_Vector', ioConfig='TCPIP0::127.0.0.1::hislip0::INSTR', devType='Source', ch=1):
        """Adds a modulated source in the Modulate tab in Modulation Distortion.
        
        Args:
            sourceName (str): User-specified name of the source to be added. [default is 'VXT']
            driver (str): Instrument driver to be used to control the source. [see validDrivers below, default is 'VXT_Vector']
            ioConfig (str): VISA address of the source. [default is x]
            devType (str): Specifies what kind of device is being added. [default (and really only option) is 'Source']
            ch (int): Channel for which source is added. [default is 1]
        """

        # print(self.inst.query(f'system:configure:edevice:cat?'))
        exists = int(self.inst.query(f'system:configure:edevice:exists? "{sourceName}"'))
        if exists == 1:
            print(f'Source "{sourceName}" already exists.')
        else:
            self.inst.write(f'system:configure:edevice:add "{sourceName}"')
            
            validDrivers = ['AGESG', 'AGEXG', 'AGGeneric', 'Agile Vector Adapter', 'AGMXG', 'AGPSG', 'M8190', 'M8190 + IQ Mixer', 'MXG_Vector', 'PSG_Vector', 'VXG', 'VXT_Vector', 'M9383A', 'Hybrid Source']
            if driver not in validDrivers:
                raise ValueError(f"Invalid driver: {driver}, must be {validDrivers}")
            self.inst.write(f'system:configure:edevice:driver "{sourceName}", "{driver}"')
            
            # Device type should ALWAYS be 'Source', but adding in others for generic purposes
            validDevTypes = ['Source', 'Power Meter', 'DC Source', 'Pulse Generator', 'SMU']
            if devType not in validDevTypes:
                raise ValueError("Invalid 'devType', must be 'Source', 'Power Meter', 'DC Source', 'Pulse Generator', or 'SMU'.")
            self.inst.write(f'system:configure:edevice:dtype "{sourceName}", "{devType}"')

            self.inst.write(f'system:configure:edevice:ioconfig "{sourceName}", "{ioConfig}"')
            self.err_check()
            self.inst.write(f'system:configure:edevice:source:modulation:control:state "{sourceName}", 1')
            
            # print(self.inst.query(f'SYSTem:CONFigure:EDEVice:CAT?'))
            # print(self.inst.query(f'sense{ch}:role:cat?'))
            # This line throws an error that the sense:mixer commands should be used instead
            # self.inst.write(f'sense{ch}:role:device "INPUT", "{sourceName}"')
            # self.inst.write(f'SENSe{ch}:DISTortion:PATH:DUT:INPut ')

        self.inst.write(f'system:configure:edevice:ioenable "{sourceName}", 1')
        
        # this is the 'Active Show in UI' checkbox
        self.inst.write(f'system:configure:edevice:state "{sourceName}", 1')
        self.wait_for_opc()
        
        self.err_check()

    def configure_mod_modulate(self, sourceName='VXT', modFile='C:\\Users\\Instrument\\Desktop\\example.mdx', enableMod=1, enableSrcCorr=0, enableLoFeedthru=0, reset=0, ch=1):
        """Configures modulation source, file, and source correction state.
        
        Args:
            sourceName (str): User-specified name of the source to be used. [default is 'VXT']
            modFile (str): Full absolute location of .mdx file to be loaded. [default is x]
            enableMod (int): Enables or disables modulated signal playback. [0, 1, default is 1]
            enableSrcCorr (int): Enables or disables source correction in the loaded file. [0, 1, default is 0]
            enableLoFeedThru (int): Enables or disables LO feedthrough correction in the loaded file. [0, 1, default is 0]
            ch (int): Channel for which settings are configured. [default is 1]            
        """

        # print(self.inst.query(f'source:cat?'))

        if reset:
            self.inst.write(f'source{ch}:modulation:file:initialize "{sourceName}"')

        self.inst.write(f'sense{ch}:distortion:modulate:source "{sourceName}"')

        # I guess source:modulation:FILE:load doesn't work but source:modulation:load does work.
        self.inst.write(f'source{ch}:modulation:load "{modFile}", "{sourceName}"')
        # print(self.inst.query(f'source{ch}:modulation:file? "{sourceName}"'))
        
        self.inst.write(f'source{ch}:modulation:correction:state {enableSrcCorr}')
        self.inst.write(f'source{ch}:modulation:correction:collection:lo:fthru:enable {enableLoFeedthru}')

        # Here I go writing bad code again
        self.inst.write(f'source{ch}:modulation:state {enableMod}, "{sourceName}"')
        self.wait_for_opc()

        self.err_check()

    def configure_mod_create_mtone_mdx(self, span=100e6, toneSpacing=100e3, numTones=1001, carrOffset=0, phaseType='random', dacScaling=70, enableOptimizer=0, fileName='C:/Users/Public/Documents/Network Analyzer/multitone.mdx', ch=1):
        """Creates and saves a multitone stimulus mdx file.
        
        Args:
            span (float): Multitone signal span in Hz. [default is x]
            toneSpacing (float): Spacing in Hz between the tones in multitone signal. [default is x]
            numTones (int): Number of tones in multitone signal. [default is 1001]
            carrOffset (float): Offset in Hz from specified carrier frequency. [default is 0]
            phaseType (str): Phase relationship between tones in multitone signal. ['random', 'fixed', 'parabolic', default is 'random']
            dacScaling (int): Unitless scaling factor used to ensure DAC filter doesn't output a signal larger than maximum output level. Setting to 100 will usually cause excessive distortion [1-100, default is 70]
            enableOptimizer (int): Determines whether to enable signal optimizer, which improves the harmonics and images of signal. [0, 1, default is 0]
            fileName (str): Full absolute path to the file that will contain the generated multitone waveform. [default is x]
            ch (int): Channel for which the waveform is created. [default is 1]
        """

        self.inst.write(f'source{ch}:modulation:file:type flattones')
        
        self.inst.write(f'source{ch}:modulation:file:signal:span:priority 1')
        self.inst.write(f'source{ch}:modulation:file:signal:span {span}')
        self.inst.write(f'source{ch}:modulation:file:signal:tone:spacing:priority 1')
        self.inst.write(f'source{ch}:modulation:file:signal:tone:spacing {toneSpacing}')
        self.inst.write(f'source{ch}:modulation:file:signal:tone:number:priority 1')
        self.inst.write(f'source{ch}:modulation:file:signal:tone:number {numTones}')
        self.inst.write(f'source{ch}:modulation:file:signal:carrier:offset {carrOffset}')
        self.inst.write(f'source{ch}:modulation:file:signal:phase:type {phaseType}')
        self.inst.write(f'source{ch}:modulation:file:signal:dac:scaling {dacScaling}')
        self.inst.write(f'source{ch}:modulation:file:signal:optimize:enable {enableOptimizer}')

        self.inst.write(f'source{ch}:modulation:file:save "{fileName}"')

    def configure_mod_create_compact_mdx(self, originalFileName='C:/Users/Public/Documents/Network Analyzer/iq.csv', sampleRate=0, span=0, toneSpacing=0, numTones=0, carrOffset=0, dacScaling=70, enableOptimizer=1, fileName='C:/Users/Public/Documents/Network Analyzer/compact.mdx', ch=1):
        """Creates and saves a compact waveform mdx file.
        
        Args:
            originalFileName (str): Full absolute path to the file that contains the original waveform to be compacted. [default is x]
            sampleRate (float): Sample rate in samples/sec of original waveform file. [default is x]
            toneSpacing (float): Tone spacing in Hz of the compact signal. [default is x]
            numTones (int): Number of tones in the compact signal. [default is x]
            carrOffset (float): Offset in Hz from specified carrier frequency. [default is 0]
            dacScaling (int): Unitless scaling factor used to ensure DAC filter doesn't output a signal larger than maximum output level. Setting to 100 will usually cause excessive distortion [1-100, default is 70]
            enableOptimizer (int): Determines whether to enable signal optimizer, which improves the harmonics and images of signal. [0, 1, default is 1]
            fileName (str): Full absolute path to the file that will contain the generated compact waveform. [default is x]
            ch (int): Channel for which the waveform is created. [default is 1]
        """
        
        self.inst.write(f'source{ch}:modulation:file:type compact')
        
        self.inst.write(f'source{ch}:modulation:file:signal:compact:ofile "{originalFileName}"')
        
        """Do these commands even need to be added??"""
        #self.inst.write(f'source{ch}:modulation:file:signal:compact:ofile:srate {sampleRate}')
        #self.inst.write(f'source{ch}:modulation:file:signal:span:priority 1')
        #self.inst.write(f'source{ch}:modulation:file:signal:span {span}')
        #self.inst.write(f'source{ch}:modulation:file:signal:tone:spacing:priority 1')
        #self.inst.write(f'source{ch}:modulation:file:signal:tone:spacing {toneSpacing}')
        #self.inst.write(f'source{ch}:modulation:file:signal:tone:number:priority 1')
        #self.inst.write(f'source{ch}:modulation:file:signal:tone:number {numTones}')
        #self.inst.write(f'source{ch}:modulation:file:signal:carrier:offset {carrOffset}')
        #self.inst.write(f'source{ch}:modulation:file:signal:dac:scaling {dacScaling}')
        #self.inst.write(f'source{ch}:modulation:file:signal:optimize:enable {enableOptimizer}')
        
        self.inst.write(f'source{ch}:modulation:file:save "{fileName}"')
    
    def configure_mod_source_cal(self, calType='power', maxIterations=3, desiredTolerance=0.1, calSpan=100e6, guardBand=0, port=1, enableCal=1, ch=1):
        """
        Configures a single calibration type in the Source Cal menu in the Modualate tab in Modulation Distortion.
        !!! If multiple calibration types are desired (for examle equalization and power), this method must be used once for each calibration type. 
        
        Args:
            calType (str): Type of calibration to perform. ['power', 'equalization', 'distortion', 'acp', default is 'power']
            maxIterations (int): Maximum number of iterations to perform correction. [default is 3]
            desiredTolerance (float): Desired correction level in dB/dB-pk-pk/dBc for power/equalization/acp. [default is 0.1]
            calSpan (float): Span in Hz over which to perform calibration. [default is x]
            guardBand (float): Frequency gap in Hz between main and adjacent channel calibrations. [default is 0]
            port (int): DUT port at which calibration will be performed. [1, 2, default is 1]
            enableCal (int): Enables or disables a given calibration row. [0, 1, default is 1]
            ch (int): Channel for which settings are configured. [default is 1]
        """
        
        validCalTypes = ['power', 'equalization', 'acp', 'distortion']
        if calType not in validCalTypes:
            raise ValueError("Invalid 'calType', must be 'power', 'equalization', or 'acp'.")
        
        validPorts = [1, 2]
        if port not in validPorts:
            raise ValueError("Invalid 'port', must be 1 or 2.")
        
        if calType == 'acp':
            self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:lower:enable {enableCal}')
            self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:upper:enable {enableCal}')
        else:
            self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:enable {enableCal}')
        
        if port == 1:
            if calType == 'acp':
                self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:lower:receiver "DUTIn1"')
                self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:upper:receiver "DUTIn1"')
            else:
                self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:receiver "DUTIn1"')
        else:
            if calType == 'acp':
                self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:lower:receiver "DUTOut2"')
                self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:upper:receiver "DUTOut2"')
            else:
                self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:receiver "DUTOut2"')
        
        if calType == 'acp':
            # these 4 aren't working properly yet
            self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:lower:span {calSpan}')
            self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:upper:span {calSpan}')
            self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:lower:gband {guardBand}')
            self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:upper:gband {guardBand}')
            self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:lower:iterations {maxIterations}')
            self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:upper:iterations {maxIterations}')
            self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:lower:tolerance {desiredTolerance}')
            self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:upper:tolerance {desiredTolerance}')
        else:
            self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:span {calSpan}')
        
        self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:iterations {maxIterations}')
        self.inst.write(f'source{ch}:modulation{port}:correction:collection:{calType}:tolerance {desiredTolerance}')
        self.wait_for_opc()

        self.err_check()
    
    def configure_mod_source_cal_details(self, enableFastCal=0, rfPowerType='fixed', power=-30, rfCarrierType='fixed', freq=1e9, stopPower=-20, stopFreq=2e9, powerPoints=11, freqPoints=11, ch=1):
        """Configures everything in the Mod Cal Details submenu within the Source Cal menu in the Modulate tab in Modulation Distortion. Wow.
        
        Args:
            enableFastCal (int): Enables or disables fast cal with reduced accuracy. [0, 1, default is 0]
            rfPowerType (str): Chooses type of power sweep over which cal is performed. ['fixed', 'swept', default is 'fixed']
            power (float): Power in dBm at which the calibration is performed. Also used as the start power if power is swept. [default is x]
            rfCarrierType (str): Chooses type of frequency sweep over which the cal is performed. ['fixed', 'swept', defaults is 'fixed']
            freq (float): Frequency in Hz at which the calibration is performed. Also used as the start frequency if frequency is swept. [default is x]
            stopPower (float): Stop power in dBm used to define the end of a power sweep. Unused if power type is 'fixed' [default is x]
            stopFrequency (float): Stop frequency in Hz used to define the end of a frequency sweep. Unused if frequency type is 'fixed' [default is x]
            powerPoints (int): Number of points to include in the power sweep. Unused if power type is 'fixed' [default is x] 
            freqPoints (int): Number of points to include in the frequency sweep. Unused if frequency type is 'fixed' [default is x]
            ch (int): Channel for which the source correction details are configured. [default is 1]
        """
        
        self.inst.write(f'source{ch}:modulation:correction:collection:fast:enable {enableFastCal}')
        
        validRfPowerTypes = ['fixed', 'swept']
        if rfPowerType.lower() not in validRfPowerTypes:
            raise ValueError("Invalid 'rfPowerType', must be 'fixed' or 'swept'.")
        
        validRfCarrierTypes = ['fixed', 'swept']
        if rfCarrierType.lower() not in validRfCarrierTypes:
            raise ValueError("Invalid 'rfCarrierType', must be 'fixed' or 'swept'.")
        
        self.inst.write(f'source{ch}:modulation:correction:collection:power:type {rfPowerType}')
        if rfPowerType.lower() == 'fixed':
            # doesn't work
            self.inst.write(f'source{ch}:modulation:correction:collection:power:fixed {power}')
        else:
            # why do these not work but the frequency ones do???
            self.inst.write(f'source{ch}:modulation:correction:collection:power:start {power}')
            self.inst.write(f'source{ch}:modulation:correction:collection:power:stop {stopPower}')
            self.inst.write(f'source{ch}:modulation:correction:collection:power:points {powerPoints}')
        
        self.inst.write(f'source{ch}:modulation:correction:collection:frequency:type {rfPowerType}')
        if rfCarrierType.lower() == 'fixed':
            # doesn't work
            self.inst.write(f'source{ch}:modulation:correction:collection:frequency:fixed {freq}')
        else:
            self.inst.write(f'source{ch}:modulation:correction:collection:frequency:start {freq}')
            self.inst.write(f'source{ch}:modulation:correction:collection:frequency:stop {stopFreq}')
            self.inst.write(f'source{ch}:modulation:correction:collection:frequency:points {freqPoints}')
    

    # def initiate_source_correction_cal(self, timeoutMs=60000, sourceName='VXT', check=1, ch=1):
    #     """Initiates a source correction calibration and turns on the correction.
        
    #     Args:
    #         timeoutMs (int): Temporary timeout for acquiring the calibration in milliseconds. [default is 60000]
    #         sourceName (str): User-specified name of the source to be used. [default is 'VXT']            
    #         check (int): If 1, rerun the source correction until it is successful. If 0, run once regardless of success or failure. [default is 1]
    #         ch (int): Channel for which the source correction is acquired. [default is 1]

    #     Returns:
    #         calDetails (list(str)): Verbose details about correction process.
    #     """

    #     calDetails = []

    #     if check:
    #         successFlag = 0
    #         while successFlag == 0:
    #             self.inst.write(f'source{ch}:modulation:correction:collection:acquire synchronous')
    #             self.wait_for_opc(tempTimeout=timeoutMs)
    #             calDetails.append(self.inst.query(f'SOUR{ch}:MOD:CORR:COLL:ACQ:DETails? "{sourceName}"').strip())
    #             calStatus = self.inst.query(f'sour{ch}:mod:corr:coll:acq:status? "{sourceName}"').strip()
                
    #             # Take a new acquisition to check the input ACP value
    #             self.single_trigger(ch)
    #             acp = self.get_mod_data('ACP LoIn1 dBc')

    #             # Only set the successFlag to 1 if both ACP is good enough and calibration is successful
    #             if acp < -33 and "Calibration succeeded." in calStatus:
    #                 successFlag = 1
    #     else:
    #         self.inst.write(f'source{ch}:modulation:correction:collection:acquire synchronous')
    #         self.wait_for_opc(tempTimeout=timeoutMs)
    #         calDetails.append(self.inst.query(f'SOUR{ch}:MOD:CORR:COLL:ACQ:DETails? "{sourceName}"').strip())

    #     return calDetails

    def initiate_source_correction_cal(self, timeoutMs=60000, sourceName='VXT', retry=1, maxRetries=3, ch=1):
        """Initiates a source correction calibration and turns on the correction.
        
        Args:
            timeoutMs (int): Temporary timeout for acquiring the calibration in milliseconds. [default is 60000]
            sourceName (str): User-specified name of the source to be used. [default is 'VXT']            
            retry (int): If 1, rerun the source correction until it is successful. If 0, run once regardless of success or failure. [default is 1]
            maxRetries (int): Specifies the maximum number of retries if the source correction fails. [default is 3]
            ch (int): Channel for which the source correction is acquired. [default is 1]
        """
        
        if retry:
            successFlag = 0
            numRetries = 0
            while successFlag == 0 and numRetries <= maxRetries:
                self.inst.write(f'source{ch}:modulation:correction:collection:acquire synchronous')
                self.wait_for_opc(tempTimeout=timeoutMs)
                calStatus = self.inst.query(f'sour{ch}:mod:corr:coll:acq:status? "{sourceName}"').strip()
                if "Calibration succeeded." in calStatus:
                    successFlag = 1
                else:
                    numRetries += 1
        else:
            self.inst.write(f'source{ch}:modulation:correction:collection:acquire synchronous')
            self.wait_for_opc(tempTimeout=timeoutMs)

        
    def enable_source_correction(self, correctionType='modpwr', enable=1, ch=1):
        """Enables or disables an existing source correction calibration.
        
        Args:
            correctionType (str): Type of correction to apply. ['modpwr', 'power', 'modulation', default is 'modpwr']
            enable (int): Enables or disables specified correction. [0, 1, default is 1]
            ch (int): Channel for which correction is specified. [default is 1]
        """
        
        if correctionType not in ['modpwr', 'power', 'modulation']:
            raise ValueError("Incorrect 'correctionType' value, must be 'modpwr', 'power', or 'modulation'")

        if correctionType == 'modpwr' and enable:
            self.inst.write(f'source{ch}:correction:select modpwr')
        elif correctionType == 'power' and enable:
            self.inst.write(f'source{ch}:correction:select power')
        elif correctionType == 'modulation' and enable:
            self.inst.write(f'source{ch}:correction:select modulation')
        else:
            self.inst.write(f'source{ch}:correction:select off')

    def configure_mod_evm_meas(self, offsetFreq=0, integBw=100e6, autofill=1, ch=1):
        """Configures the Measure tab as an EVM measurement in Modulation Distortion.
        
        Args:
            offsetFreq (float): Offset of the modulated signal in Hz from the carrier frequency. [default is 0]
            integBw (float): Bandwidth of the modulated signal in Hz. [default is x]
            autoFill (int): Specifies whether to autofill all parameters based on information in .mdx file. [0, 1, default is 1]
            ch (int): Channel in which the modulation measurement will be configured. [default is 1]
        """
        
        # Sets the measurement type
        self.inst.write(f'sense{ch}:distortion:measure:band:type EVM')
        
        if autofill:
            self.inst.write(f'sense{ch}:distortion:measure:band:autofill')
        else:
            self.inst.write(f'sense{ch}:distortion:measure:band:carrier:offset {offsetFreq}')
            self.inst.write(f'sense{ch}:distortion:measure:band:carrier:ibw {integBw}')

        self.wait_for_opc()
        self.err_check()    

    def configure_mod_acpevm_meas(self, offsetFreq=0, integBw=100e6, acpLoFreq=-100e6, acpLoIntegBw=100e6, acpUpFreq=100e6, acpUpIntegBw=100e6, autofill=1, ch=1):
        """Configures the Measure tab as an ACP and EVM measurement in Modulation Distortion.
        
        Args:
            offsetFreq (float): Offset of the modulated signal in Hz from the carrier frequency. [default is 0]
            integBw (float): Bandwidth of the modulated signal in Hz. [default is x]
            acpLoFreq (float): Offset in Hz from the center of the main channel to the center of the low side adjacent channel. [default is x]
            acpLoIntegBw (float): Bandwidth of the low side adjacent channel. [default is x]
            acpUpFreq (float): Offset in Hz from the center of the main channel to the center of the high side adjacent channel. [default is x]
            acpUpIntegBw (float): Bandwidth of the high side adjacent channel. [default is x]
            autoFill (int): Specifies whether to autofill all parameters based on information in .mdx file. [0, 1, default is 1]
            ch (int): Channel in which the modulation measurement will be configured. [default is 1]
        """
        
        # Sets the measurement type
        self.inst.write(f'sense{ch}:distortion:measure:band:type ACPEVM')
        
        if autofill:
            self.inst.write(f'sense{ch}:distortion:measure:band:autofill')
        else:
            self.inst.write(f'sense{ch}:distortion:measure:band:carrier:offset {offsetFreq}')
            self.inst.write(f'sense{ch}:distortion:measure:band:carrier:ibw {integBw}')
            self.inst.write(f'sense{ch}:distortion:measure:band:acp:lower:offset {acpLoFreq}')
            self.inst.write(f'sense{ch}:distortion:measure:band:acp:lower:ibw {acpLoIntegBw}')
            self.inst.write(f'sense{ch}:distortion:measure:band:acp:upper:offset {acpUpFreq}')
            self.inst.write(f'sense{ch}:distortion:measure:band:acp:upper:ibw {acpUpIntegBw}')
        
        self.wait_for_opc()
        self.err_check()

    def configure_mod_meas_details(self, eqApertureAuto=1, eqAperture=15e6, antiAliasFilter='auto', evmNorm=1.0, modFilter='none', modFilterAlpha=0.35, symRateAuto=1, symRate=200e6, dutNf=0, ch=1):
        """Configures the Measurement Details dialog under the Measure tab in Modulation Distortion.
        
        Args:
            eqApertureAuto (int): Determine whether to automatically set the distortion aperture window size. [0, 1, default is 1]
            eqAperture (float): Window span in Hz used to model the DUT's gain and distortion. Unused if eqApertureAuto is 1. [default is x]
            antiAliasFilter (str): Select the anti aliasing filter used to remove aliasing in the distortion tone measurements. ['auto', 'wide', 'narrow', default is 'auto']
            evmNorm (float): Sets the scaling factor applied to the EVM measurements. [0.1-1.0, default is 1.0]
            modFilter (str): Chooses pulse shaping filter applied to the modulated signal measurement. ['none', 'rrc', default is 'none']
            modFilterAlpha (float): Unitless rolloff factor for pulse shaping filter. [0.1-1.0, default is x]
            symRateAuto (int): Determine whether to automatically choose pulse shaping filter symbol rate. [0, 1, default is 1]
            symRate (float): Sets the symbol rate in sym/sec of pulse shaping filter. Ignored if symRateAuto is 1. [default is x]
            dutNf (float): Sets the estimated noise figure in dB of DUT to be included in distortion measurement. [default is x]
        """
        
        self.inst.write(f'sense{ch}:distortion:measure:correlation:aperture:auto {eqApertureAuto}')
        
        if not eqApertureAuto:
            self.inst.write(f'sense{ch}:distortion:measure:correlation:aperture {eqAperture}')
            
        validAliasFilters = ['narrow', 'wide', 'auto']
        if antiAliasFilter.lower() not in validAliasFilters:
            raise ValueError("Invalid 'antiAliasFilter', must be 'narrow', 'wide', or 'auto'.")
        
        self.inst.write(f'sense{ch}:distortion:adc:filter:type {antiAliasFilter}')
        self.inst.write(f'sense{ch}:distortion:evm:normalize {evmNorm}')
        
        validModFilters = ['none', 'rrc']
        if modFilter.lower() not in validModFilters:
            raise ValueError("Invalid 'modFilter', must be 'none' or 'rrc'.")
        
        self.inst.write(f'sense{ch}:distortion:measure:filter {modFilter}')
        
        if modFilter == 'rrc':
            self.inst.write(f'sense{ch}:distortion:measure:filter:alpha {modFilterAlpha}')
            self.inst.write(f'sense{ch}:distortion:measure:filter:srate:auto {symRateAuto}')
            if not symRateAuto:
                self.inst.write(f'sense{ch}:distortion:measure:filter:srate {symRate}')
        
        self.inst.write(f'sense{ch}:distortion:path:dut:nominal:nf {dutNf}')
    
    def add_mod_table_parameter(self, modTableParam='Carrier In1 dBm', ch=1):
        """Adds the specified parameter to the Modulation Distortion table.
        
        Args:
            modTableParam (str): Parameter to be added to the modulation distortion table. [see validParams, default is 'Carrier In1 dBm']
            ch (int): Channel to which the parameter will be added. [default is 1]
        """
        
        validParams = ['Carrier In1 dBm', 'Carrier Out2 dBm', 'Carrier Gain21 dB', 'Carrier IBW',
                        'Carrier OffsFreq', 'EVM DistEq21 dBc', 'EVM DistEq21 %', 'EVM InEq1 dBc',
                        'EVM InEq1 %', 'EVM OutEq2 dBc', 'EVM OutEq2 %', 'ACP LoIn1 dBc',
                        'ACP LoOut2 dBc', 'ACP UpIn1 dBc', 'ACP UpOut2 dBc',
                        'Carrier In2 dBm', 'Carrier Out1 dBm', 'Carrier Gain12 dB',
                        'EVM DistEq12 dBc', 'EVM DistEq12 %', 'EVM InEq2 dBc',
                        'EVM InEq2 %', 'EVM OutEq1 dBc', 'EVM OutEq1 %', 'ACP LoIn2 dBc',
                        'ACP LoOut1 dBc', 'ACP UpIn2 dBc', 'ACP UpOut1 dBc']
        
        if modTableParam not in validParams:
            raise ValueError("Invalid 'modTableParam', check measurement parameter argument.")

        self.inst.write(f'sense{ch}:distortion:table:display:feed "{modTableParam}"')
        self.wait_for_opc()
    
    def delete_mod_table_parameter(self, modTableParam='Carrier In1 dBm', ch=1):
        """Deletes the specified parameter from the Modulation Distortion table.
        
        Args:
            modTableParam (str): Parameter to be deleted from the modulation distortion table. [see validParams, default is 'Carrier In1 dBm']
            ch (int): Channel from which the parameter will be deleted. [default is 1]
        """
        
        validParams = ['Carrier In1 dBm', 'Carrier Out2 dBm', 'Carrier Gain21 dB', 'Carrier IBW', 
                        'Carrier OffsFreq', 'EVM DistEq21 dBc', 'EVM DistEq21 %', 'EVM InEq1 dBc', 
                        'EVM InEq1 %', 'EVM OutEq2 dBc', 'EVM OutEq2 %', 'ACP LoIn1 dBc', 
                        'ACP LoOut2 dBc', 'ACP UpIn1 dBc', 'ACP UpOut2 dBc',
                        'Carrier In2 dBm', 'Carrier Out1 dBm', 'Carrier Gain12 dB', 
                        'EVM DistEq12 dBc', 'EVM DistEq12 %', 'EVM InEq2 dBc', 
                        'EVM InEq2 %', 'EVM OutEq1 dBc', 'EVM OutEq1 %', 'ACP LoIn2 dBc', 
                        'ACP LoOut1 dBc', 'ACP UpIn2 dBc', 'ACP UpOut1 dBc']
        if modTableParam not in validParams:
            raise ValueError("Invalid 'modTableParam', check measurement parameter argument.")

        self.inst.write(f'sense{ch}:distortion:table:display:delete "{modTableParam}"')
        self.wait_for_opc()
    
    def save_mod_distortion_table(self, fileName='C:/Users/Public/Documents/Network Analyzer/distortion_table.csv', ch=1):
        """Saves the data in the Modulation Distortion table to a csv file on the VNA PC.
        
        Args:
            fileName (str): Location to save the distortion table csv file. [default is x]
            ch (int): Channel from which the distortion table will be saved. [default is 1]
        """
        
        self.inst.write(f'sense{ch}:distortion:table:display:save "{fileName}"')
        self.wait_for_opc()
    
    def show_distortion_table(self, win=1, enabled=1):
        """Shows or hides the Modulation Distortion table.
        
        Args:
            win (int): Window in which the distortion table is configured. [default is 1]
            enabled (int): Specifies whether to hide or show the distortoin table. [0, 1, default is 1]
        """
        
        if enabled:
            self.inst.write(f'display:window{win}:table distortion')
            self.wait_for_opc()
        else:
            self.inst.write(f'display:window{win}:table off')
            self.wait_for_opc()
    
    def get_mod_data(self, modTableParam='Carrier In1 dBm', ch=1):
        """Gets a specific value from the Modulation Distortion table.
        
        Args:
            modTableParam (str): Parameter to be extracted from the modulation distortion table. [see validParams, default is 'Carrier In1 dBm']
            ch (int): Channel from which the parameter is extracted. [default is 1]
        """
        
        validParams = ['Carrier In1 dBm', 'Carrier Out2 dBm', 'Carrier Gain21 dB', 'Carrier IBW', 
                        'Carrier OffsFreq', 'EVM DistEq21 dBc', 'EVM DistEq21 %', 'EVM InEq1 dBc', 
                        'EVM InEq1 %', 'EVM OutEq2 dBc', 'EVM OutEq2 %', 'ACP LoIn1 dBc', 
                        'ACP LoOut2 dBc', 'ACP UpIn1 dBc', 'ACP UpOut2 dBc',
                        'Carrier In2 dBm', 'Carrier Out1 dBm', 'Carrier Gain12 dB', 
                        'EVM DistEq12 dBc', 'EVM DistEq12 %', 'EVM InEq2 dBc', 
                        'EVM InEq2 %', 'EVM OutEq1 dBc', 'EVM OutEq1 %', 'ACP LoIn2 dBc', 
                        'ACP LoOut1 dBc', 'ACP UpIn2 dBc', 'ACP UpOut1 dBc']
        
        if modTableParam not in validParams:
            raise ValueError(f"Invalid 'modTableParam' {modTableParam}, check measurement parameter argument.")
        
        # This should always be 1 since there will only ever be one measurement band specified at a time
        bandNumber = 1
        
        self.inst.write('format ascii') 
        raw = self.inst.query(f'sense{ch}:distortion:table:data:value? {bandNumber},"{modTableParam}"')
        return float(raw.rstrip())
        # return raw.rstrip()
        
    # endregion
    
    # region Modulation Distortion Converters
    def configure_modx_rfpath(self, srcAtten=0, includeSrcAtten=1, inputPort=1, outputPort=2, recvAtten=0, ch=1):
        """Configures the RF Path tab in Modulation Distortion Converters.
        
        Args:
            srcAtten (int): Source attenuation in dB. [0, 10, 20, 30, 40, default is 0]
            includeSrcAtten (int): Specifies whether to include source attenuation in power calculations. [0, 1, default is 1]
            inputPort (int): Specifies which VNA port corresponds to DUT input port. [1, 2, default is 1]
            outputPort (int): Specifies which VNA port corresponds to DUT output port. [1, 2, default is 2]
            recvAtten (int): Receiver attenuation in dB. [0, 10, 20, 30, 40, default is 0]
        """
        
        validPorts = [1, 2]
        if inputPort not in validPorts or outputPort not in validPorts:
            raise ValueError("Invalid 'inputPort' or 'outputPort', must be 1 or 2.")
        
        # There's an issue with the source attenuation command if the source port is not 1
        self.inst.write(f'source{ch}:power{inputPort}:attenuation {srcAtten}')

        self.inst.write(f'sense{ch}:distortion:path:source:attenuation:include {includeSrcAtten}')
        
        self.inst.write(f'SENSe{ch}:DISTortion:PATH:DUT:INPut {inputPort}')
        self.inst.write(f'SENSe{ch}:DISTortion:PATH:DUT:OUTPut {outputPort}')
        # self.inst.write(f'sense{ch}:mixer:pmap {inputPort},{outputPort}')
        self.inst.write(f'source{ch}:power{outputPort}:attenuation:receiver:test {recvAtten}')

        self.err_check()

    def new_modx_trace(self, measName='InputPower', measParam='PIn1', modify=0, win=1, ch=1):
        """Creates a new trace in a Modulation Distortion Converters channel.
        
        Args:
            measName (str): Name of the trace to be added. [default is 'InputPower']
            measParam (str): Specifies the trace to be added. [see validParams, default is 'PIn1']
            modify (int): Specifies if a new measurement will be created or an existing measurement will be modified. [0, 1, default is 0]
            win (int): Window to which trace will be added. [default is 1]
            ch (int): Channel to which trace will be added. [default is 1]
        """
        
        validParams = ['PIn1', 'POut1', 'POut2', 'PModFile', 'MSig2', 'MDist2', 'MDistIR2', 'MGain21', 
                        'PGain21', 'MComp21', 'PGain21', 'LMatch2', 'CarrIn1', 'CarrOut2', 'NPRIn1', 'NPROut2', 
                        'NPRDist21', 'NPRPwrOut2', 'ACPIn1', 'ACPOut2', 'ACPDist21', 'ACPPwrIn1', 'ACPPwrOut2', 
                        'EVMDistEq21', 'EVMDistUn21', 'EVMPwrIn1', 'EVMPwrOut2', 'ModFilter', 
                        'A', 'b1', 'B', 'C', 'b3', 'D', 'b4', 'R1', 'a1', 'R2', 'a2', 'R3', 'a3', 'R4', 'a4',
                        'S11', 'S21', 'LPIn1', 'LPOut1', 'LPOut2',
                        'PIn2', 'POut2', 'POut1', 'MSig1', 'MDist1', 'MDistIR1', 'MGain12', 
                        'PGain12', 'MComp12', 'PGain12', 'LMatch1', 'CarrIn2', 'CarrOut1', 'NPRIn2', 'NPROut1', 
                        'NPRDist12', 'NPRPwrOut1', 'ACPIn2', 'ACPOut1', 'ACPDist12', 'ACPPwrIn2', 'ACPPwrOut1', 
                        'EVMDistEq12', 'EVMDistUn12', 'EVMPwrIn2', 'EVMPwrOut1']
        
        if measParam not in validParams:
            raise ValueError("Invalid 'measParam', check measurement parameter argument.")

        if modify:
            # Select a measurement and change its measurement parameter
            self.inst.write(f'calc{ch}:parameter:select "{measName}"')
            self.inst.write(f'calc{ch}:custom:modify "{measParam}"')
            self.err_check()
        else:
            # Create new trace with name, parameter, and channel
            self.inst.write(f'calc{ch}:custom:define "{measName}", "Modulation Distortion Converters", "{measParam}"')
            self.err_check()
            
            # Check for windows
            windows = self.inst.query('display:catalog?').split(',')
            
            # The VNA returns *everything* as a string, so to compare the desired window to the available windows in the VNA, we have to convert the 'win' argument to a string
            # 1 in ['1', '2'] would return False, but '1' in ['1', '2'] will return True, which is what we want.
            
            # If the specified window doesn't exist, create it by turning it on
            if str(win) not in windows:
                self.inst.write(f'display:window{win}:state on')

            # See how many traces are in the window
            traces = self.inst.query(f"display:window{win}:catalog?").strip()

            # Determine the number of the 'next' trace in the window
            # VNA returns the string 'EMPTY' if a window has no traces defined in it
            if 'empty' in traces.lower():
                nextTrace = 1
            else:
                nextTrace = len(traces.split(',')) + 1

            # Feed the new trace into a window and check for errors
            self.inst.write(f'display:window{win}:trace{nextTrace}:feed "{measName}"')
        
        self.wait_for_opc()
        self.err_check()
    
    def configure_modx_mixer(self, loFreq=1e9, sideband='low', ch=1):
        """Sets the LO frequency and mixer sideband for mixer measurements in Modulation Distortion Converters.
        
        Args:
            loFreq (float): LO frequency of the DUT in Hz. [default is x]
            sideband (str): Selects which mixer sideband the VNA will use for measurement. ['low', 'high', default is 'low']
            ch (int): Channel for which settings are configured. [default is 1]
        """
        
        inputFreq = self.inst.query(f'sense{ch}:distortion:sweep:carrier:frequency?')

        self.inst.write(f'sense{ch}:mixer:input:frequency:fixed {inputFreq}')
        self.inst.write(f'sense{ch}:mixer:lo1:frequency:fixed {loFreq}')
        # self.inst.write(f'sense{ch}:mixer:calculate output')
        self.wait_for_opc()

        validSidebands = ['low', 'high']
        if sideband.lower() not in validSidebands:
            raise ValueError("Invalid 'sideband', must be 'low' or 'high'.")
        
        self.inst.write(f'sense{ch}:mixer:output:frequency:sideband {sideband}')
        self.inst.write(f'sense{ch}:mixer:calculate output')
        self.wait_for_opc()
   
    def configure_modx_embedded_lo(self, tuningMethod='broadband', sweepInterval=1, span=3e6, noiseBw=10e3, iterations=5, tolerance=5, enable=1, ch=1):
        """Configures the embedded LO dialog for mixer measurements in Modulation Distortion Converters.
        
        Args:
            tuningMethod (str): Chooses tuning method for embedded LO search. ['broadband', 'precise', 'none', default is 'broadband']
            sweepInterval (int): Chooses how often to search for the LO in sweeps. [default is 1 sweep between each LO search]
            span (float): Span in Hz for broadband LO search. [use a value that is at least as wide as the expected offset between expected and actual LO frequency, default is 3 MHz]
            noiseBw (float): Receiver IF bandwidth in Hz used for the LO search. [lower value reduces noise floor and slows down sweep, default is 10 kHz]
            iterations (int): Maximum number of iterations to find LO. [default is 5]
            tolerance (float): Minimum frequency offset in Hz between previous and current LO measurements for a stable measurement. [default is 5 Hz]
            enable (int): Enables or disables embedded LO. [0, 1, default is 1]
            ch (int): Channel for which embedded LO is configured. [default is 1]
        """        
        
        validTuningMethods = ['broadband', 'precise', 'none']
        if tuningMethod not in validTuningMethods:
            raise ValueError("Invalid 'tuningMethod', must be 'broadband', 'precise', or 'none'.")
        
        self.inst.write(f'sense{ch}:mixer:elo:tuning:mode {tuningMethod}')
        self.inst.write(f'sense{ch}:mixer:elo:tuning:interval {sweepInterval}')
        self.inst.write(f'sense{ch}:mixer:elo:tuning:span {span}')
        self.inst.write(f'sense{ch}:mixer:elo:tuning:nbw {noiseBw}')
        self.inst.write(f'sense{ch}:mixer:elo:tuning:iterations {iterations}')
        self.inst.write(f'sense{ch}:mixer:elo:tuning:tolerance {tolerance}')
        
        time.sleep(1)
        
        self.inst.write(f'sense{ch}:mixer:elo:state {enable}')

    # endregion
    
    # region Gain Compression
    def new_gca_trace(self, measName='GainAtCompression', measParam='CompGain21', modify=0, win=1, ch=1):
        """Creates a new trace in a Gain Compression channel.
        
        Args:
            measName (str): Name of trace/measurement, e.g. "MyVeryCoolS21Measurement". [default is 'GainAtCompression']
            measParam (str): Name of parameter to be measured by trace, e.g. "S21". [see validParams, default is 'CompGain21']
            modify (int): Specifies whether to modify an existing measurement or create a new one. [0=create new trace, 1=modify existing trace, default is 0]
            win (int): Window where the trace/measurement will be displayed. [default is 1]
            ch (int): Channel to which the trace/measurement will be assigned. [default is 1]
        """

        validParams = ['S21', 'S11', 'S12', 'S22', 'CompIn21', 'CompOut21', 'DeltaGain21', 'CompGain21', 'CompS11', 'RefS21', 'CompIn12', 'CompOut12', 'DeltaGain12', 'CompGain12', 'CompS22', 'RefS12']
        if measParam not in validParams:
            raise ValueError(f"Invalid 'measParam' {measParam}, check measurement parameter argument.")
        
        if modify:
            # Select a measurement and change its measurement parameter
            self.inst.write(f'calc{ch}:parameter:select "{measName}"')
            self.inst.write(f'calc{ch}:custom:modify "{measParam}"')
            self.err_check()
        else:
            # Create new trace with name, parameter, and channel
            self.inst.write(f'calc{ch}:custom:define "{measName}", "Gain Compression", "{measParam}"')
            
            # Check for windows
            windows = self.inst.query('display:catalog?').split(',')
            
            # The VNA returns *everything* as a string, so to compare the desired window to the available windows in the VNA, we have to convert the 'win' argument to a string
            # 1 in ['1', '2'] would return False, but '1' in ['1', '2'] will return True, which is what we want.
            
            # If the specified window doesn't exist, create it by turning it on
            if str(win) not in windows:
                self.inst.write(f'display:window{win}:state on')
            
            # See how many traces are in the window
            traces = self.inst.query(f"display:window{win}:catalog?").strip()
            
            # Determine the number of the 'next' trace in the window
            # VNA returns the string 'EMPTY' if a window has no traces defined in it
            if 'empty' in traces.lower():
                nextTrace = 1
            else:
                nextTrace = len(traces.split(',')) + 1
            
            # Feed the new trace into a window and check for errors
            self.inst.write(f'display:window{win}:trace{nextTrace}:feed "{measName}"')
            self.wait_for_opc()
    
    def configure_gca_frequency_stimulus(self, sweepType='linear', acqMode='smartsweep', numPoints=201, startFreq=1e9, stopFreq=2e9, ifBw=100e3, ch=1):
        """Configure the settings in the frequency tab in Gain Compression setup.
        
        Args:
            sweepType (str): Selects the type of stimulus sweep to use. 'linear' and 'logarithmic are with respect to frequency. ['linear', 'logarithmic', 'power', 'cw', 'phase', default is 'linear']
            acqMode (str): Selects the acquisition mode used to capture power and frequency data. Smartsweep is significantly faster than the others. ['smartsweep', 'pfrequency', 'fpower', default is 'smartsweep']
            numPoints (int): Number of sweep points to be used. [default is 201]
            startFreq (float): Start frequency in Hz. [default is x]
            stopFreq (float): Stop frequency in Hz. [default is x]
            ifBw (float): IF bandwidth in Hz. [default is x]
            ch (int): Channel for which settings are configured. [default is 1]
        """
        
        # Error checking
        validSweepTypes = ['linear', 'logarithmic', 'power', 'cw', 'phase']
        if sweepType.lower() not in validSweepTypes:
            raise ValueError("Invalid 'sweepType', must be 'linear', 'logarithmic', 'power', 'cw', or 'phase'.")
        
        validAcqModes = ['pfrequency', 'fpower', 'smartsweep']
        if acqMode.lower() not in validAcqModes:
            raise ValueError("Invalid 'acqMode', must be 'pfrequency', 'fpower', or 'smartsweep'.")
        
        # Configure GCA settings - sweep
        self.inst.write(f'sense{ch}:sweep:type {sweepType}')
        self.inst.write(f'sense{ch}:gcsetup:amode {acqMode}')
        self.inst.write(f'sense{ch}:gcsetup:sweep:frequency:points {numPoints}')
        self.inst.write(f'sense{ch}:frequency:start {startFreq}')
        self.inst.write(f'sense{ch}:frequency:stop {stopFreq}')
        self.inst.write(f'sense{ch}:bandwidth {ifBw}')
        
        self.wait_for_opc()
    
    def configure_gca_power_stimulus(self, inputPort=1, outputPort=2, linearPower=-30, reversePower=-20, startPower=-30, stopPower=-10, ch=1):
        """Configure the settings in the power tab in Gain Compression Setup.
        
        Args:
            inputPort (int): Selects which VNA port is connected to the DUT input. [1, 2, default is 1]
            outputPort (int): Selects which VNA port is connected to the DUT output. [1, 2, default is 2]
            linearPower (float): Sets the forward sweep power in dBm to be used to calculate the linear reference gain. [default is x]
            reversePower (float): Sets the referse sweep power in dBm to be used to calculate the reverse s-parameters. [default is x]
            startPower (float): Sets the start power in dBm used for the power sweep. [default is x]
            stopPower (float): Sets the stop power in dBm used for the power sweep. [default is x]
            ch (int): Channel for which settings are configured. [default is 1]
        """

        validPorts = [1, 2]
        if inputPort not in validPorts or outputPort not in validPorts:
            raise ValueError("Invalid 'inputPort' or 'outputPort', must be 1 or 2.")

        # Configure GCA settings - power
        self.inst.write(f'sense{ch}:gcsetup:pmap {inputPort},{outputPort}')
        self.inst.write(f'sense{ch}:gcsetup:power:linear:input:level {linearPower}')
        self.inst.write(f'sense{ch}:gcsetup:power:reverse:level {reversePower}')
        self.inst.write(f'sense{ch}:gcsetup:power:start:level {startPower}')
        self.inst.write(f'sense{ch}:gcsetup:power:stop:level {stopPower}')

        self.wait_for_opc()

    def configure_gca_safe_mode_stimulus(self, safeMode=0, coarseInc=3, fineInc=1, thresh=0.5, limit=20, ch=1):
        """Configure the SMART Sweep Safe Mode settings in the compression tab in Gain Compression Setup.
        
        Args:
            safeMode (int): Turns Safe Mode on or off. [0, 1, default is 0]
            coarseInc (float): Sets the coarse power increment in dB that Safe Mode will use to increase the power. [default is 3 dB]
            fineInc (float): Sets the fine power increment in dB that Safe Mode will use to increase the power when within "threshold" of the compression point. [default is 1 dB]
            thresh (float): Sets the threshold in dB from the desired compression point at which Safe Mode will adjust the power from coarseInc to fineInc. [default is 0.5 dB]
            limit (float): Sets the maximum power in dBm that Safe Mode will allow the VNA source to be set to. [default is 20 dBm]
            ch (int): Channel for which settings are configured. [default is 1]
        """

        # Configure GCA settings - Safe Mode
        if safeMode:
            self.inst.write(f'sense{ch}:GCSetup:SAFE:ENABle {safeMode}')
            self.inst.write(f'sense{ch}:GCSetup:SAFE:CPADjustment {coarseInc}')
            self.inst.write(f'sense{ch}:GCSetup:SAFE:FPADjustment {fineInc}')
            self.inst.write(f'sense{ch}:GCSetup:SAFE:FTHReshold {thresh}')
            self.inst.write(f'sense{ch}:GCSetup:SAFE:MLimit {limit}')
        else:
            self.inst.write(f'sense{ch}:GCSetup:SAFE:ENABle {safeMode}')

        self.wait_for_opc()

    def configure_gca_compression_analysis(self, measName, cwFreq, ch=1):
        """Configures compression analysis (underlying power sweep in GCA).

        Args:
            measName (string): Specifies the measurement name.
            cwFreq (float): Specifies cw frequency.
            ch (int): Specifies the channel of the measurement. [default is 1]
        """
        
        # Select trace and get trace number
        self.inst.write(f'calculate{ch}:parameter:select "{measName}"')
        measNum = int(self.inst.query(f'calculate{ch}:parameter:mnumber?'))

        # Configure compression analysis
        self.inst.write(f'calculate{ch}:measure{measNum}:gcmeas:analysis:enable 1')
        self.inst.write(f'calculate{ch}:measure{measNum}:gcmeas:analysis:discrete:state 0')
        self.inst.write(f'calculate{ch}:measure{measNum}:gcmeas:analysis:cwfrequency {cwFreq}')
        self.inst.write(f'calculate{ch}:measure{measNum}:gcmeas:analysis:xaxis pin')
        
        self.wait_for_opc()
    
    def get_cw_freq(self, measName, ch=1):
        """Gets the CW frequency associated with the selected trace.

        Args:
            measName (str): Measurement from which the CW frequency will be retrieved.
            ch (int): Channel that the selected measurement is on. [default is 1]

        Returns:
            cwFreq (float): CW frequency of the trace.
        """

        # Select trace
        self.inst.write(f'calculate{ch}:parameter:select "{measName}"')
        
        # Get measurement number of the selected trace (we need to use this to query the status of compression analysis)
        measNum = int(self.inst.query(f'calculate{ch}:parameter:mnumber?'))
        
        # Determine if compression analysis is turned on or off (only applies to GCA measurement class)
        if 'Gain Compression' in self.inst.query('system:active:mclass?'):
            compStatus = self.inst.query(f'calculate{ch}:measure{measNum}:gcmeas:analysis:enable?').strip()
            
            # If compression analysis is on, get CW frequency, otherwise, return the Python keyword 'None'
            if compStatus == '1':
                cwFreq = float(self.inst.query(f'calculate{ch}:measure{measNum}:gcmeas:analysis:cwfrequency?'))
            else:
                cwFreq = None
        # If non-GCA measurement class, just return the Python keyword 'None'
        else:
            cwFreq = None
            
        self.wait_for_opc()

        return cwFreq
    
    # endregion
    
    # region Gain Compression Converters
    def new_gcax_trace(self, measName='GainAtCompression', measParam='CompGain21', win=1, ch=1):
        """Creates a new trace in a Gain Compression Converters channel.
        
        measName (str): Name of trace/measurement, e.g. "MyVeryCoolS21Measurement". [default is 'GainAtCompression']
            measParam (str): Name of parameter to be measured by trace, e.g. "S21". [default is 'CompGain21']
            win (int): Window where the trace/measurement will be displayed. [default is 1]
            ch (int): Channel to which the trace/measurement will be assigned. [default is 1]
        """
        
        validParams = ['S21', 'S11', 'S12', 'S22', 'CompIn21', 'CompOut21', 'DeltaGain21', 'CompGain21', 'CompS11', 'RefS21',
                       'SC21', 'SC12', 'Ipwr', 'RevIPwr', 'Opwr', 'RevOPwr']
        if measParam not in validParams:
            raise ValueError("Invalid 'measParam', check measurement parameter argument.")
        
        # Create new trace with name, parameter, and channel
        self.inst.write(f'calc{ch}:custom:define "{measName}", "Gain Compression Converters", "{measParam}"')
        
        # Check for windows
        windows = self.inst.query('display:catalog?').split(',')
        
        # The VNA returns *everything* as a string, so to compare the desired window to the available windows in the VNA, we have to convert the 'win' argument to a string
        # 1 in ['1', '2'] would return False, but '1' in ['1', '2'] will return True, which is what we want.
        
        # If the specified window doesn't exist, create it by turning it on
        if str(win) not in windows:
            self.inst.write(f'display:window{win}:state on')
        
        # See how many traces are in the window
        traces = self.inst.query(f"display:window{win}:catalog?").strip()
        
        # Determine the number of the 'next' trace in the window
        # VNA returns the string 'EMPTY' if a window has no traces defined in it
        if 'empty' in traces.lower():
            nextTrace = 1
        else:
            nextTrace = len(traces.split(',')) + 1
        
        # Feed the new trace into a window and check for errors
        self.inst.write(f'display:window{win}:trace{nextTrace}:feed "{measName}"')
        self.wait_for_opc()
    
    def configure_gcax_frequency_stimulus(self, sweepType='linear', acqMode='smartsweep', numPoints=201, startFreq=1e9, stopFreq=2e9, ifBw=100e3, ch=1):
        """Passthrough method for clarity."""

        self.configure_gca_frequency_stimulus(sweepType, acqMode, numPoints, startFreq, stopFreq, ifBw, ch)

    def configure_gcax_power_stimulus(self, inputPort=1, outputPort=2, linearPower=-30, reversePower=-20, startPower=-30, stopPower=-10, ch=1):
        """Passthrough method for clarity."""

        self.configure_gca_power_stimulus(inputPort, outputPort, linearPower, reversePower, startPower, stopPower, ch)

    def configure_gcax_safe_mode_stimulus(self, safeMode=0, coarseInc=3, fineInc=1, thresh=0.5, limit=20, ch=1):
        """Passthrough method for clarity."""

        self.configure_gca_safe_mode_stimulus(safeMode, coarseInc, fineInc, thresh, limit, ch)

    #endregion

    #region Generic Mixers
    def configure_mixer_frequency(self, startFreq=3e9, stopFreq=3.5e9, loFreq=2.5e9, sideband='low', inputGreater=1, ch=1):
        """
        Configures input start/stop frequency, LO fixed frequency, and output sideband
        and calculates output start/stop frequency in Gain Compression Converters and 
        Scalar Mixer Converter channels.

        Args:
            startFreq (float): Start frequency of DUT input. [default is x]
            stopFreq (float): Stop frequency of DUT input. [default is x]
            sideband (str): Selects which converter sideband to use. ['low', 'high', default is 'low']
            inputGreater (int): Tells VNA if DUT is a downconverter (input freq greater than LO) or an upconverter (input freq less than LO). [0, 1, default is 1]
            ch (int): Channel for which settings are configured. [default is 1]
        """
        
        self.inst.write(f'sense{ch}:mixer:input:frequency:start {startFreq}')
        self.inst.write(f'sense{ch}:mixer:input:frequency:stop {stopFreq}')
        self.inst.write(f'sense{ch}:mixer:lo:frequency:fixed {loFreq}')
        self.inst.write(f'sense{ch}:mixer:lo:frequency:ilti {inputGreater}')
        
        validSidebands = ['low', 'high']
        if sideband.lower() not in validSidebands:
            raise ValueError("Invalid 'sideband', must be 'low' or 'high'.")
        
        self.inst.write(f'sense{ch}:mixer:output:frequency:sideband {sideband}')
        self.inst.write(f'sense{ch}:mixer:calculate output')
        self.wait_for_opc()
    
    def configure_embedded_lo(self, tuningMethod='broadband', tuningPoint=401, sweepInterval=1, span=1e6, ifBw=10e3, iterations=5, tolerance=5, enable=1, ch=1):
        """Configures the Embedded LO dialog for mixer measurements in Gain Compression Converters and Scalar Mixer Converter channels.

        Args:
            tuningMethod (str): Chooses tuning method for embedded LO search. ['broadband', 'precise', 'none']
            sweepInterval (int): Chooses how often to search for the LO in sweeps. [default is 1 sweep between each LO search]
            tuningPoint (int): Selects which trace point is used to search for actual LO signal. [choose the center point, you'll have to know how many trace points you have]
            span (float): Span in Hz for broadband LO search. [use a value that is at least as wide as the expected offset between expected and actual LO frequency, default is 1 MHz]
            ifBw (float): Receiver IF bandwidth in Hz used for the LO search. [lower value reduces noise floor and slows down sweep, default is 10 kHz]
            iterations (int): Maximum number of iterations to find LO. [default is 5]
            tolerance (float): Minimum frequency offset in Hz between previous and current LO measurements for a stable measurement. [default is 5 Hz]
            ch (int): Channel for which settings are configured. [default is 1]
        """        
        
        self.inst.write(f'sense{ch}:mixer:elo:state {enable}')
        
        validTuningMethods = ['broadband', 'precise', 'none']
        if tuningMethod not in validTuningMethods:
            raise ValueError("Invalid 'tuningMethod', must be 'broadband', 'precise', or 'none'.")
        
        self.inst.write(f'sense{ch}:mixer:elo:tuning:mode {tuningMethod}')
        self.inst.write(f'sense{ch}:mixer:elo:normalize:point {tuningPoint}')
        self.inst.write(f'sense{ch}:mixer:elo:tuning:interval {sweepInterval}')
        self.inst.write(f'sense{ch}:mixer:elo:tuning:span {span}')
        self.inst.write(f'sense{ch}:mixer:elo:tuning:ifbw {ifBw}')
        self.inst.write(f'sense{ch}:mixer:elo:tuning:iterations {iterations}')
        self.inst.write(f'sense{ch}:mixer:elo:tuning:tolerance {tolerance}')
        time.sleep(1)
    
    def get_lo_frequency_delta(self, ch=1):
        """Gets the LO Frequency Delta in the Embedded LO dialog for mixer measurements in Modulation Distortion Converters.
        
        Args:
            ch (int): Channel from which embedded LO frequency is queried. [default is 1]
        """
        
        return float(self.inst.query(f'sense{ch}:mixer:elo:lo:delta?'))

    # endregion
    
    # region Spectrum Analyzer
    def new_sa_trace(self, measName='Spectrum', measParam='B', win=1, ch=1):
        """Creates a new trace in a spectrum analyzer channel.
        
        Args:
            measName (str): Name of trace/measurement, e.g. "MyVeryCoolS21Measurement". [default is 'Spectrum']
            measParam (str): Name of parameter to be measured by trace, e.g. "S21". [see validParams, default is 'B']
            win (int): Window where the trace/measurement will be displayed. [default is 1]
            ch (int): Channel to which the trace/measurement will be assigned. [default is 1]
        """
        
        validParams = ['B', 'A', 'R1', 'R2', 'b1', 'b2', 'a1', 'a2']
        if measParam not in validParams:
            raise ValueError("Invalid 'measParam', check measurement parameter argument.")

        # Create new trace with name, parameter, and channel
        self.inst.write(f'calc{ch}:custom:define "{measName}", "Spectrum Analyzer", "{measParam}"')
        
        # Check for windows
        windows = self.inst.query('display:catalog?').split(',')
        
        # The VNA returns *everything* as a string, so to compare the desired window to the available windows in the VNA, we have to convert the 'win' argument to a string
        # 1 in ['1', '2'] would return False, but '1' in ['1', '2'] will return True, which is what we want.
        
        # If the specified window doesn't exist, create it by turning it on
        if str(win) not in windows:
            self.inst.write(f'display:window{win}:state on')

        # See how many traces are in the window
        traces = self.inst.query(f"display:window{win}:catalog?").strip()

        # Determine the number of the 'next' trace in the window
        # VNA returns the string 'EMPTY' if a window has no traces defined in it
        if 'empty' in traces.lower():
            nextTrace = 1
        else:
            nextTrace = len(traces.split(',')) + 1

        # Feed the new trace into a window and check for errors
        self.inst.write(f'display:window{win}:trace{nextTrace}:feed "{measName}"')
        
        self.wait_for_opc()
        self.err_check()
    
    def configure_sa_sweep(self, centerFreq=1e9, span=1e9, startFreq=500e6, stopFreq=1.5e9, numPoints=801, resBw=1e6, videoBw=1e6, detectorType='peak', 
                           useStartStop=1, resBwAuto=1, videoBwAuto=1, detectorBypass=0, ch=1):
        """Configures the SA tab in spectrum analzyer setup.

        Args:
            centerFreq (float): Center frequency in Hz. Unused if useStartStop=1. [default is x]
            span (float): Span frequency in Hz. Unused if useStartStop=1. [default is x]
            startFreq (float): Start frequency in Hz. Unused if useStartStop=0. [default is x]
            stopFreq (float): Stop frequency in Hz. Unused if useStartStop=0. [default is x]
            numPoints (int): Number of trace points. [default is 801]
            resBw (float): Resolution bandwidth in Hz. Unused if resBwAuto=1. [default is x]
            videoBw (float): Video bandwidth in Hz. Unused if videoBwAuto=1. [default is x]
            detectorType (str): Spectrum analyzer detector type. [see validDetectorTypes, default is 'peak']
            useStartStop (int): Determines whether to use start/stop [1] or center/span [0] for frequency settings. [0, 1, default is 1]
            resBwAuto (int): Determines whether to use automatic resolution bandwidth setting based on span. [0, 1, default is 1]
            videoBwAuto (int): Determines whether to use automatic video bandwidth setting based on span. [0, 1, default is 1]
            detectorBypass (int): Determines whether to bypass detector and use all FFT points. [0, 1, default is 0]
            ch (int): Channel for which settings are configured. [default is 1]
        """
        
        if useStartStop:
            self.inst.write(f'sense{ch}:frequency:start {startFreq}')
            self.inst.write(f'sense{ch}:frequency:stop {stopFreq}')
        else:
            self.inst.write(f'sense{ch}:frequency:center {centerFreq}')
            self.inst.write(f'sense{ch}:frequency:span {span}')
        
        
        self.inst.write(f'sense{ch}:sweep:points {numPoints}')
        
        if not resBwAuto:
            self.inst.write(f'sense{ch}:sa:bandwidth:resolution {resBw}')
        
        if not videoBwAuto:
            self.inst.write(f'sense{ch}:sa:bandwidth:video {videoBw}')
        
        self.inst.write(f'sense{ch}:sa:detector:bypass:state {detectorBypass}')
        
        validDetectorTypes = ['peak', 'average', 'sample', 'normal', 'negpeak', 'psample', 'paverage', 'faspeak']
        if detectorType.lower() not in validDetectorTypes:
            raise ValueError("Invalid 'detectorType', must be 'peak', 'average', 'sample', 'normal', 'negpeak', 'psample', 'paverage', or 'faspeak'.")
        
        if not detectorBypass:
            self.inst.write(f'sense{ch}:sa:detector:function {detectorType}')
            
        # self.inst.write('initiate:continuous 0')
        self.inst.write(f'sense{ch}:sweep:mode hold')
    
    def configure_sa_source(self, sourcePort=1, portStateOn=0, sourceFreq=1e9, sourcePower=-15, ch=1):
        """Configures the Source tab in spectrum analzyer setup.
        
        Args:
            sourcePort (int): Source port to be configured. [1, 2, default is 1]
            portStateOn (int): Turns source port off or on. [0, 1, default is 0]
            sourceFreq (int): CW frequency of source. [default is x]
            sourcePower (float): CW power of source. [default is x]
            ch (int): Channel for which settings are configured. [default is 1]
        """
        
        validPorts = [1, 2]
        if sourcePort not in validPorts:
            raise ValueError("Invalid 'sourcePort', must be 1 or 2.")
        
        portState = 'ON' if portStateOn else 'OFF'
        self.inst.write(f'source{ch}:power{sourcePort}:mode {portState}')
        self.inst.write(f'sense{ch}:sa:source{sourcePort}:frequency:cw {sourceFreq}')
        self.inst.write(f'sense{ch}:sa:source{sourcePort}:power:value {sourcePower}')
    
    def configure_sa_band_power_marker(self, mkrNum, measName, mkrCenterFreq=1e9, span=1e6, bandPowerState=1, ch=1):
        """Configures marker for band power SA analysis.
        
        Args:
            mkrNum (int): Marker number. [0-15]
            measName (str): Name of the measurement for which the marker will be configured.
            mkrCenterFreq (float): Marker center frequency in Hz. [default is x]
            span (float): Span in Hz over which to measure band power. [default is x]
            bandPowerState (int): Turns band power calculation off or on. [0, 1, default is 1]
            ch (int): Channel for which settings are configured. [default is 1]
        """
        
        self.marker_activate(mkrNum, measName, ch)
        measNum = self.get_meas_number_from_name(measName, ch)
        self.inst.write(f'calculate{ch}:measure{measNum}:sa:marker{mkrNum}:bpower:state {bandPowerState}')
        self.marker_set_x(mkrNum, measName, mkrCenterFreq, ch)
        self.inst.write(f'calculate{ch}:measure{measNum}:sa:marker{mkrNum}:bpower:span {span}')
        # wait_for_opc doesn't work for the above command.
        # sleep as a workaround for now so that the analyzer can adjust band power span
        # before the band power is queried.
        time.sleep(1)
        
    def get_sa_marker_band_power(self, mkrNum, measName, ch=1):
        """Gets band power level from band power marker.
        
        Args:
            mkrNum (int): Marker number. [0-15]
            measName (str): Name of the measurement to which the marker will be added.
            ch (int): Channel for which settings are configured. [default is 1]
        """
        
        # Certain SCPI commands use measurement number instead of measurement name
        # You can get one from the other, thus this helper function
        measNum = self.get_meas_number_from_name(measName, ch)
        
        raw = self.inst.query(f'calculate{ch}:measure{measNum}:sa:marker{mkrNum}:bpower:data?')
        return float(raw.rstrip())
    
    # endregion
    
    # region Scalar Mixer Converters
    def new_smc_trace(self, measName='ForwardConversion', measParam='SC21', win=1, ch=1):
        """Creates a new trace in a Scalar Mixer Converters channel.
        
        Args:
            measName (str): Name of trace/measurement, e.g. "MyVeryCoolS21Measurement". [default is 'ForwardConversion']
            measParam (str): Name of parameter to be measured by trace, e.g. "S21". [default is 'SC21']
            win (int): Window where the trace/measurement will be displayed. [default is 1]
            ch (int): Channel to which the trace/measurement will be assigned. [default is 1]
        """
        
        validParams = ['SC21', 'SC12', 'S11', 'S22', 'Ipwr', 'RevIPwr', 'Opwr', 'RevOPwr']
        if measParam not in validParams:
            raise ValueError("Invalid 'measParam', check measurement parameter argument.")
        
        # Create new trace with name, parameter, and channel
        self.inst.write(f'calc{ch}:custom:define "{measName}", "Scalar Mixer/Converter", "{measParam}"')
        
        # Check for windows
        windows = self.inst.query('display:catalog?').split(',')
        
        # The VNA returns *everything* as a string, so to compare the desired window to the available windows in the VNA, we have to convert the 'win' argument to a string
        # 1 in ['1', '2'] would return False, but '1' in ['1', '2'] will return True, which is what we want.
        
        # If the specified window doesn't exist, create it by turning it on
        if str(win) not in windows:
            self.inst.write(f'display:window{win}:state on')
        
        # See how many traces are in the window
        traces = self.inst.query(f"display:window{win}:catalog?").strip()
        
        # Determine the number of the 'next' trace in the window
        # VNA returns the string 'EMPTY' if a window has no traces defined in it
        if 'empty' in traces.lower():
            nextTrace = 1
        else:
            nextTrace = len(traces.split(',')) + 1
        
        # Feed the new trace into a window and check for errors
        self.inst.write(f'display:window{win}:trace{nextTrace}:feed "{measName}"')
        self.wait_for_opc()
    
    def configure_smc_stimulus(self, inputPort=1, outputPort=2, portPower=-15, numPoints=201, ifBw=10e3, ch=1):
        """Configure the ports, power, number of points, and IF bandwidth for an SMC measurement.
        
        Arguments:
            inputPort (int): Port connected to DUT input. [1, 2, default is 1]
            outputPort (int): Port connected to DUT output. [1, 2, default is 2]
            portPower (float): VNA output power in dBm. [default is x]
            numPoints (int): Number of points to measure. [default is 201]
            ifBw (int): IF Bandwidth in Hz. [default is x]
            ch (int): Channel for which stimulus is configured. [default is 1]
        """
        
        self.inst.write(f'sense{ch}:mixer:pmap {inputPort},{outputPort}')
        self.inst.write(f'source{ch}:power {portPower}')
        self.inst.write(f'sense{ch}:sweep:points {numPoints}')
        self.inst.write(f'sense{ch}:bandwidth {ifBw}')

        self.inst.write('initiate:continuous 0')

        self.wait_for_opc()
    
    # endregion
    
    # region Noise Figure
    def new_nf_trace(self, measName='Noise Figure', measParam='NF', modify=0, win=1, ch=1):
        """Creates a new trace in a Noise Figure channel.
        
        Args:
            measName (str): Name of trace/measurement, e.g. "MyVeryCoolS21Measurement". [default is 'Noise Figure']
            measParam (str): Name of parameter to be measured by trace, e.g. "S21". [see validParams, default is 'NF']
            modify (int): 0 creates a new measurement, 1 modifies an existing measurement. [default is 0]
            win (int): Window where the trace/measurement will be displayed. [default is 1]
            ch (int): Channel to which the trace/measurement will be assigned. [default is 1]
        """

        validParams = ['NF', 'ENR', 'T-Eff', 'DUTRNP', 'DUTRNPI', 'SYSRNP', 'SYSRNPI', 'DUTNPD', 'DUTNPDI', 'SYSNPD', 'SYSNPDI', 'OvrRng', 'T-Rcvr', 'S11', 'S21', 'S12', 'S22', 'GammaOpt', 'Rn', 'NFMin']
        if measParam not in validParams:
            raise ValueError(f"Invalid 'measParam' {measParam}, check measurement parameter argument.")
        
        if modify:
            # Select a measurement and change its measurement parameter
            self.inst.write(f'calc{ch}:parameter:select "{measName}"')
            self.inst.write(f'calc{ch}:custom:modify "{measParam}"')
            self.err_check()
        else:
            # Create new trace with name, parameter, and channel
            self.inst.write(f'calc{ch}:custom:define "{measName}", "Noise Figure Cold Source", "{measParam}"')
            
            # Check for windows
            windows = self.inst.query('display:catalog?').split(',')
            
            # The VNA returns *everything* as a string, so to compare the desired window to the available windows in the VNA, we have to convert the 'win' argument to a string
            # 1 in ['1', '2'] would return False, but '1' in ['1', '2'] will return True, which is what we want.
            
            # If the specified window doesn't exist, create it by turning it on
            if str(win) not in windows:
                self.inst.write(f'display:window{win}:state on')
            
            # See how many traces are in the window
            traces = self.inst.query(f"display:window{win}:catalog?").strip()
            
            # Determine the number of the 'next' trace in the window
            # VNA returns the string 'EMPTY' if a window has no traces defined in it
            if 'empty' in traces.lower():
                nextTrace = 1
            else:
                nextTrace = len(traces.split(',')) + 1
            
            # Feed the new trace into a window and check for errors
            self.inst.write(f'display:window{win}:trace{nextTrace}:feed "{measName}"')
            self.wait_for_opc()
        self.err_check()

    def configure_nf_frequency(self, sweepType='linear', numPoints=41, startFreq=10e6, stopFreq=44e9, ifBw=1e3, ch=1):
        """Configure the settings in the frequency tab in Noise Figure setup.
        
        Args:
            sweepType (str): Selects the type of stimulus sweep to use. 'linear' and 'logarithmic are with respect to frequency. ['linear', 'logarithmic', 'power', 'cw', 'phase', default is 'linear']
            numPoints (int): Number of sweep points to be used. [default is 411]
            startFreq (float): Start frequency in Hz. [default is 10e6]
            stopFreq (float): Stop frequency in Hz. [default is 44e9]
            ifBw (float): IF bandwidth in Hz. [default is 1e3]
            ch (int): Channel for which settings are configured. [default is 1]
        """
        
        validSweepTypes = ['linear', 'logarithmic', 'power', 'cw', 'phase']
        if sweepType.lower() not in validSweepTypes:
            raise ValueError("Invalid 'sweepType', must be 'linear', 'logarithmic', 'power', 'cw', or 'phase'.")
        
        self.inst.write(f'SENSe{ch}:FREQuency:STARt {startFreq}')
        self.inst.write(f'SENSe{ch}:FREQuency:STOP {stopFreq}')
        self.inst.write(f'SENSe{ch}:SWEep:TYPE {sweepType}')
        self.inst.write(f'SENSe{ch}:SWEep:POINts {numPoints}')
        self.inst.write(f'SENSe{ch}:BWIDth:RESolution {ifBw}')
        self.err_check()

    def configure_nf_power(self, inputPort=1, outputPort=2, power=-20, reversePower=-20, ch=1):
        """Configure the settings in the power tab in Noise Figure Setup.
        
        Args:
            inputPort (int): Selects which VNA port is connected to the DUT input. [1, 2, default is 1]
            outputPort (int): Selects which VNA port is connected to the DUT output. [1, 2, default is 2]
            power (float): Sets the forward sweep power in dBm to be used to calculate the DUT gain. [default is x]
            reversePower (float): Sets the referse sweep power in dBm to be used to calculate the reverse s-parameters. [default is x]
            ch (int): Channel for which settings are configured. [default is 1]
        """

        validPorts = [1, 2]
        if inputPort not in validPorts or outputPort not in validPorts:
            raise ValueError("Invalid 'inputPort' or 'outputPort', must be 1 or 2.")
        
        # self.inst.write(f'OUTPut:STATe 1')
        self.inst.write(f'SENSe{ch}:NOISe:PMAP {inputPort},{outputPort}')
        self.inst.write(f'SOURce{ch}:POWer{inputPort}:LEVel:IMMediate:AMPLitude {power}')
        self.inst.write(f'SOURce{ch}:POWer{outputPort}:LEVel:IMMediate:AMPLitude {reversePower}')
        self.err_check()

    def configure_nf_noise_figure(self, noiseBw=4e6, avgState=1, avgNum=100, receiverGain=30, sourceTemp=297, ch=1):
        """Configure the settings in the noise figure tab in Noise Figure setup.

        Args:
            noiseBw (float): Bandwidth of the noise receiver. [default is 4e6]
            avgState (int): 0 turns averaging off. 1 turns averaging on. [default is 1]
            avgNum (int): Number of averages. [default is 100]
            receiverGain (int): Sets noise receiver gain. [0, 15, 30, default is 30]
            sourceTemp (int): Ambient temperature at which the noise measurement is occurring. [default is 297]
        """
        
        validNoiseBw = [800e3, 2e6, 4e6, 8e6, 24e6]
        if noiseBw not in validNoiseBw:
            raise ValueError(f'Invalid "noiseBw": {noiseBw}, must be one of {validNoiseBw}')
        self.inst.write(f'SENSe{ch}:NOISe:BWIDth:RESolution {noiseBw}')

        self.inst.write(f'SENSe{ch}:NOISe:AVERage:COUNt {avgNum}')
        self.inst.write(f'SENSe{ch}:NOISe:AVERage:STATe {avgState}')
        
        if receiverGain not in [0, 15, 30]:
            raise ValueError(f'Invalid "receiverGain": {receiverGain}, must be 0, 15, or 30')
        self.inst.write(f'SENSe{ch}:NOISe:GAIN {receiverGain}')
        
        self.inst.write(f'SENSe{ch}:NOISe:TEMPerature:SOURce {sourceTemp}')
        self.err_check()
    # endregion

    # region Noise Figure Converters
    def new_nfx_trace(self, measName='Noise Figure', measParam='NF', modify=0, win=1, ch=1):
        """Creates a new trace in a Noise Figure channel.
        
        Args:
            measName (str): Name of trace/measurement, e.g. "MyVeryCoolS21Measurement". [default is 'Noise Figure']
            measParam (str): Name of parameter to be measured by trace, e.g. "S21". [see validParams, default is 'NF']
            modify (int): 0 creates a new measurement, 1 modifies an existing measurement. [default is 0]
            win (int): Window where the trace/measurement will be displayed. [default is 1]
            ch (int): Channel to which the trace/measurement will be assigned. [default is 1]
        """

        validParams = ['NF', 'ENR', 'T-Eff', 'DUTRNP', 'DUTRNPI', 'SYSRNP', 'SYSRNPI', 'DUTNPD', 'DUTNPDI', 'SYSNPD', 'SYSNPDI', 'OvrRng', 'T-Rcvr', 'S11', 'SC21', 'SC12', 'S22', 'Ipwr', 'RevIPwr', 'Opwr', 'RevOPwr']
        if measParam not in validParams:
            raise ValueError(f"Invalid 'measParam' {measParam}, check measurement parameter argument.")
        
        if modify:
            # Select a measurement and change its measurement parameter
            self.inst.write(f'calc{ch}:parameter:select "{measName}"')
            self.inst.write(f'calc{ch}:custom:modify "{measParam}"')
            self.err_check()
        else:
            # Create new trace with name, parameter, and channel
            self.inst.write(f'calc{ch}:custom:define "{measName}", "Noise Figure Converters", "{measParam}"')
            
            # Check for windows
            windows = self.inst.query('display:catalog?').split(',')
            
            # The VNA returns *everything* as a string, so to compare the desired window to the available windows in the VNA, we have to convert the 'win' argument to a string
            # 1 in ['1', '2'] would return False, but '1' in ['1', '2'] will return True, which is what we want.
            
            # If the specified window doesn't exist, create it by turning it on
            if str(win) not in windows:
                self.inst.write(f'display:window{win}:state on')
            
            # See how many traces are in the window
            traces = self.inst.query(f"display:window{win}:catalog?").strip()
            
            # Determine the number of the 'next' trace in the window
            # VNA returns the string 'EMPTY' if a window has no traces defined in it
            if 'empty' in traces.lower():
                nextTrace = 1
            else:
                nextTrace = len(traces.split(',')) + 1
            
            # Feed the new trace into a window and check for errors
            self.inst.write(f'display:window{win}:trace{nextTrace}:feed "{measName}"')
            self.wait_for_opc()
        self.err_check()
    # endregion

    # region PXIe Switch Driver
    def spdt_enable(self, state=1, module=1, ch=1):
        """Enables control of M9155 PXIe Dual SPDT switch card in the VNA firmware.
        
        Args:
            state (int): Turns control of M9155 card off or on. [0, 1, default is 1]
            module (int): Selects which M9155 module to control. [1, 2, default is 1]
            ch (int): Channel for which switch control is enabled. [default is 1]
        """

        self.inst.write(f'sense{ch}:switch:m9155:module{module}:control:state {state}')
        self.err_check()
    
    def spdt_get_path_catalog(self, module=1, switch=1, ch=1):
        """HELPER FUNCTION: Returns catalog of valid paths for the M9155 PXIe Dual SPDT switch card.
        
        Args:
            module (int): Selects which M9155 module to control. [default is 1]
            switch (int): Selects which of the two SPDTs in the M9155 to control. [1, 2, default is 1]
            ch (int): Channel for which switch control is enabled. [default is 1]
        """

        return self.inst.query(f'sense{ch}:switch:m9155:module{module}:switch{switch}:path:catalog?')

    def spdt_close_connection(self, module=1, switch=1, route=1, ch=1):
        """Closes a path in one of the SPDTs in the M9155 PXIe Dual SPDT switch card.
        
        Args:
            module (int): Selects which M9155 module to control. [default is 1]
            switch (int): Selects which of the two SPDTs in the M9155 to control. [1, 2, default is 1]
            route (int): Selects which route to close in the selected SPDT. [1, 2, default is 1]
        """
        
        validRoutes = [1, 2]
        if route not in validRoutes:
            raise ValueError(f'Invalid route {route} selected. Choose from {validRoutes}.')
        
        self.inst.write(f'sense{ch}:switch:m9155:module{module}:switch{switch}:path state{route}')

    def spdt_connection_status(self, module=1, switch=1, ch=1):
        """Gets the status of a given SPDT switch in the M9155 PXIe Dual SPDT switch card.
        
        Args:
            module (int): Selects which M9155 module to control. [default is 1]
            switch (int): Selects which of the two SPDTs in the M9155 to control. [1, 2, default is 1]
        """
        
        return self.inst.query(f'sense{ch}:switch:m9155:module{module}:switch{switch}:path?')

    def sp6t_enable(self, state=1, module=1, ch=1):
        """Enables control of M9157 PXIe SP6T switch card.
        
        Args:
            state (int): Turns control of SP6T card off or on. [0, 1, default is 1]
            module (int): Selects which M9157 card to enable. [default is 1]
            ch (int): Channel for which switch control is enabled. [default is 1]
        """

        self.inst.write(f'sense{ch}:switch:m9157:module{module}:control:state {state}')
        self.err_check()

    def sp6t_close_connection(self, module=1, route=1, ch=1):
        """Closes a path in the M9157 PXIe SP6T switch card.
        
        Args:
            module (int): Selects which M9157 card to control. [default is 1]
            route (int): Selects which route to close in the SP6T. [1-6, default is 1]
        """
        
        validRoutes = [1, 2, 3, 4, 5, 6]
        if route not in validRoutes:
            raise ValueError(f'Invalid route {route} selected. Choose from {validRoutes}.')
        
        self.inst.write(f'sense{ch}:switch:m9157:module{module}:switch:path state{route}')
    
    def sp6t_connection_status(self, module=1, ch=1):
        """Gets the status of the SP6T in the M9157 PXIe SP6T switch card.
        
        Args:
            module (int): Selects which M9157 card to control. [default is 1]
        """
        
        return self.inst.query(f'sense{ch}:switch:m9157:module{module}:switch:path?')
    
    # endregion

    # region
    """Lowest pair of region tags won't work for some reason if there isn't another pair below them."""
    # endregion

