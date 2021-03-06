# Pore-Bubble_Analysis
Fiji Macro to automate large scale analysis of foams, porous solids and bubbly liquids characteristics. Applications include polymeric foams, metallic foams, foams in foods (bread, doughnuts, ice cream, aerated chocolate, mousse), porous rocks, wood products, bioengineering structures (e.g., hydroxyapatite, CAP scaffolds) and porous catalysts. This code is most useful in the analysis of large data sets / big data image analysis.
We used this to analyse Bread cross-section images to determine 'fluffiness' as a consequence of sonication (https://doi.org/10.1088/0964-1726/18/10/104001) but this could be also used to analyse any porous structure or foamed material, as well as bubbles in liquids or emulsions.
We have chosen to upload what we have created as a contribution to the Community that has been an invaluable resource in our recent work - analysing baked goods and thir characteristics to investigate the application of ultrasonic irradiation during prooving, baking or frying and to identify optimum processing parameters. We hope that this code may be useful to others in their analysis and assessment of other food stuffs, porous materials, voids and bubbles.

Code takes an image input and creates a file of results with tables and output images (distribution plots and heat maps), characterising bubble size, orientation and global porosity.
We have made this folder of '.ijm' codes. To install this into Fiji, copy the whole unzipped folder into the plugins file of your local Fiji.app program files (i.e. C:\Program Files\Fiji.app\plugins)

Folder Contains 'Additional_Operations', a folder that contain other codes used to prepare the images and rename them in preparation to be fed to the Code. These were one-time-use codes but have been left as they may prove useful in the future.
Also, included, are excel (VBA) macros, that were used to collate and summerise results.
Please consult the provided PDF document as a Guide on how the code was used by the MMM Lab Research Team.
Input: Tiff, Jpeg or PNG images of porous medium (cropped to the area of interest and set/known scale)
Output: A folder of various results, including distribution plots and heat maps of key pore characteristics, as well as pore
	results tables and measurements of global characteristics, crust thickness and ring hole characteristics (if applicable).

Written by Ben Sargeant, for the Multifunctional Materials Manufacturing Lab Research Group ??? Loughborough University.
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
		This has not been modified and are other Free Software???s. These are also covered by the GNU General Public License
		and come with no warranty. 
    
>> MorphoLibJ (Legland, D., Arganda-Carreras, I., & Andrey, P. (2016). MorphoLibJ: integrated library and plugins for mathematical
	morphology with ImageJ. Bioinformatics, btw413. doi:10.1093/bioinformatics/btw413 )
		Used to clean images after thresholding and identify ring holes.
		This has not been modified and are other Free Software???s. These are also covered by the GNU General Public License
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
