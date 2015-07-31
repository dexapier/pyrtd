Excel RealTimeData (RTD) components were introduced in Excel 2002 providing
a robust mechanism for inserting real-time data into a spreadsheet. Behind the
scenes a RTD server is a COM object which implements the IRTDServer interface.

**pyrtd** implements a Python RTD client that can receive real-time data from any RTD component.

# Installation #
The latest release of pyrtd is always available on PyPI at http://pypi.python.org/pypi/pyrtd/.

You can install the pyrtd library from the command line using easy\_install

```
easy_install pyrtd
```

or pip

```
pip install pyrtd
```

# Example #
```
"""
Example usage of rtd.RTDClient connecting to the RTDTime.RTD component. The
RTDTime.dll implementing this component is part of the "Building a Real-Time
Data Server in Excel 2002" MSDN article and is available at:

http://download.microsoft.com/download/4/9/c/49cb54f8-63b6-4024-845b-fd2c8b0d8917/odc_xlrtdbuild.exe

You'll have to register the RTD component by executing "regsvr32 RTDTime.dll".
This component sends an UpdateNotify every second, thus this example prints out
the current time every second.

"""
import pythoncom
from rtd import RTDClient

if __name__ == '__main__':
    time = RTDClient('RTDTime.RTD')
    time.connect()
    time.register_topic('Now')

    while 1:
        # This line is critical, it tells the pythoncom subsystem to
        # handle any pending windows messages. We're waiting for an
        # UpdateNotify callback from the RTDServer; if we don't
        # check for messages we'll never be notified of pending
        # RTD updates!
        pythoncom.PumpWaitingMessages()

        if time.update():
            print time.get('Now')
```
This example is available in the repository at http://code.google.com/p/pyrtd/source/browse/examples/rtdtime.py.