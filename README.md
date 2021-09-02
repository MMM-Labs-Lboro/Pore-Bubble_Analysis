# Pore-Bubble_Analysis
Fiji Marco to automate large scale analysis of foam characteristics.
We used this to analyse Bread images to determine 'fluffiness' but could be used to analyse and porous structure or material, as well as bubbles.
We have chosen to upload what we have made as a contribution to the community that has been an invaluable resource in our recent work - analysing baked goods charictoristics to investigate the application of ultrasonic irradiation during baking and identify optimum parameters. We hope that this code may be useful to others, perhaps in analysis of other food stuffs, porous materials or bubble analysis.

Code takes an image input and creates a file of results tables and output images (distribution plots and heat maps), characterising bubble size, orientation and global porosity.
We have made this folder of '.ijm' codes. To install this into Fiji, copy the whole unzipped folder into the plugins file of your local Fiji.app program files (i.e. C:\Program Files\Fiji.app\plugins)

Folder Contains Additional_Operations that contain other codes used to prepare some images and rename images. These were one-time-use codes but have been left as they may prove useful in the future.
Also, included, are excel (VBA) macros, that were used to collate and summerise results.
Please consult the provided PDF document as a Guide on how the code was used by the MMM Research Team.
Input: Tiff, Jpeg or PNG images of porous medium (cropped to the area of interest and set/known scale)
Output: A folder of various results, including distribution plots and heat maps of key pore characteristics, as well as pore
	results tables and measurements of global characteristics, crust thickness and ring hole characteristics (if applicable).

Written by Ben Sargeant, for the Multifunctional Materials Manufacturing Research Group – Loughborough University.
	Contact the MMM Labs by email on C.Torres@lboro.ac.uk.
Copyright (C) 2021 Ben Sargeant & Carmen Torres-Sanchez. This program is free software: you can redistribute it and/or modify
	it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of
	the License, or any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
	without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public
	License for more details. You should have received a copy of the GNU General Public License along with this program (file
	name LICENSE). If not, see <https://www.gnu.org/licenses/>.

These functions utilize code written by other users. These are dependencies are in addition to the standard functions of Fiji and
	require installation before this code can be used:
>> BAR (Ferreira, T., Miura, K., Bitdeli Chef, & Eglinger, J. (2015, August 21). Scripts: Bar 1.1.6. Zenodo.
	https://doi.org/10.5281/ZENODO.28838">doi:10.5281/zenodo.28838" class="csl-entry">Ferreira, T., Miura, K., Bitdeli Chef,
	& Eglinger, J. (2015, August 21). Scripts: Bar 1.1.6. Zenodo. https://doi.org/10.5281/ZENODO.28838 )
		Used to plot distribution and colour result maps.
		This has not been modified and are other Free Software’s. These are also covered by the GNU General Public License
		and come with no warranty. 
    
>> MorphoLibJ (Legland, D., Arganda-Carreras, I., & Andrey, P. (2016). MorphoLibJ: integrated library and plugins for mathematical
	morphology with ImageJ. Bioinformatics, btw413. doi:10.1093/bioinformatics/btw413 )
		Used to clean images after thresholding and identify ring holes.
		This has not been modified and are other Free Software’s. These are also covered by the GNU General Public License
		and come with no warranty. 
    
>> Read_And_Write_Excel (Anthony Sinadinos, Brenden Kromhout, 2017)
		Used to save main results to one location.
		This plugin uses the Apache POI api, which is distributed under the terms of the Apache Licence (available from
		https://poi.apache.org/legal.html). I believe this software to be free and open source.

These codes are installed by adding the plug-ins update sites to ImageJ; Open Fiji > Help > Updates > Manage update Sites
	> Check BAR, IJPB-plugins and ResultsToExcel > Close > Apply Changes > Close > Restart Fiji 

All of these have been used in good faith and without any known breach of copywrite, as these are all believed to be free software.
	Our gratitude goes to the creators of the plug-ins (referenced above) for there hard work and their contributions to the
	coding community.
