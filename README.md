![alt tag](https://raw.githubusercontent.com/lateralblast/devices/master/devices.jpg)

DEVICES
=======

Diagram Export in Visio from CSV and Excel and other Sources

License
-------

This software is licensed as CC-BA (Creative Commons By Attrbution)

http://creativecommons.org/licenses/by/4.0/legalcode

Introduction
------------

A Powershell script for creating Visio Diagrams of DC racks and the hardware in them.

The import can be in the form of CSV or Excel (still under development).

This could also be used to automate Visio digram creations using exports from CMDBs (e.g. ServiceNow, Remedy, etc).

At the moment the script is a proof of concept. It has support for a number of vendor stencils and provides a framework to expand on.

Things to do:

- Better determination of stencils names
- Adding asset and serial information to stencils that are put on page
- Excel import support

Output
------

Example output (JPG of Visio Document):

![alt tag](https://raw.githubusercontent.com/lateralblast/devices/master/rack.jpg)

Requirements
------------

The following software is required:

- Windows OS
- Powershell
- Visio
- Visio Stencils for vendor products

If you want to clone the script and/or stencils:

- Git for Windows

Documentation
-------------

You can copy the script manually from the git repository or clone it:

```
$ git clone https://github.com/lateralblast/devices.git .
```

Stencils are put in the 'stencils' subdirectory under a 'vendor' subdirectory.

To help, I'm building a collection of Visio stencils here:

https://github.com/lateralblast/vss

This repository is getting rather large so I'd recommend you just copy the ones you need.

If you wanted to clone the entire collection:

```
$ cd devices
$ git clone https://github.com/lateralblast/vss.git stencils
```

Currently there is some support for the following vendor stencils:

- Oracle
- Dell
- IBM
- NetApp
- Pure

Support for other vendors is relatively straight forward to add, 
you need to inspect the Visio file and look at the naming standard
for front and rear views. Common naming is "Model Front" and "Model Rear".

I plan to add some code to list the stencil names and do a match to make this process easier.

Usage
-----

To run the script from the command line you may need to alter the execution policy,
by setting it globally or adding the following command line option:

```
-ExecutionPolicy ByPass
```

Getting help:

```
Y:\Code\devices>powershell -ExecutionPolicy ByPass -File devices.ps1 -help
usage: devices.ps1
--help
--version
--inputfile  FILENAME
--outputfile FILENAME
```

Example of a CSV file:

```
$ more data\servers.csv
Hostname,Vendor,Architecture,Model,Rack,Top Rack Unit,Rack Units
server1,Oracle,SPARC,M3000,A1,2,2
server2,Oracle,SPARC,M5000,A1,12,10
server3,Oracle,x86,X2-4,A1,14,3
disk3,Pure,NA,Disk shelf,A1,17,2
disk2,Pure,NA,Disk shelf,A1,19,2
disk1,Pure,NA,Disk shelf,A1,21,2
disk0,Pure,NA,Disk shelf,A1,23,2
server4,Pure,NA,FA M70,A1,26,3
server5,Dell,x86,R820,A1,28,2
server11,Oracle,SPARC,M3000,A2,2,2
server12,Oracle,SPARC,M5000,A2,12,10
server13,Oracle,x86,X2-4,A2,14,3
disk3,Pure,NA,Disk shelf,A2,17,2
disk2,Pure,NA,Disk shelf,A2,19,2
disk1,Pure,NA,Disk shelf,A2,21,2
disk0,Pure,NA,Disk shelf,A2,23,2
server14,Pure,NA,FA M70,A2,26,3
server15,Dell,x86,R820,A2,28,2
```

Importing CSV file and creating Visio diagrams:

```
Y:\Code\devices>powershell -ExecutionPolicy ByPass -File devices.ps1 -inputfile data\servers.csv
Output: Y:\Code\devices\output\rack-A1.vsd
Output: Y:\Code\devices\output\rack-A2.vsd
```
