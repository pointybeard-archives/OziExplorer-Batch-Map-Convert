# OziExplorer Batch Map Convert #

- Version: 0.91
- Date: 10th January 2010
- Github Repository: <http://github.com/pointybeard/OziExplorer-Batch-Map-Convert>
- Author: Alistair Kearney <alistair@pointybeard.com>

## Synopsis

The OziExplorer map format (.OZF) is a closed proprietary format. As such, there are no 3rd party programs that can read/write this format.

This application allows for batch conversion of .OZF files via the OziExplorer API which can then be used easily with other raster map/image applications.

Please be aware that this is in an early stage of development and likely is buggy.


## Requirements

- Visual Basic 6 (to compile)
- OziExplorer v3.95.2 or above
- OziExplorer API v1.08 or above (<http://www.oziexplorer3.com/oziapi/oziapi.html>)


## Change Log
	
	0.91
		- Hidden some UI controls. Can be exposed using the "Show Advanced Options" button
		- Check for OziExplorer API 1.08 or greater when program loads
		- Check if OziExplorer is running when program loads
		- Progress dialog will always be on top
		- Removed "Cancel" button from progress dialog. Instead cancel is achieved via the [X] close button

	0.9 alpha
		- Initial Release