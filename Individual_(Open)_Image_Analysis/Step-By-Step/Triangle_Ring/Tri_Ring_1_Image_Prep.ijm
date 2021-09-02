// Macro to automate analysis of baked goods pore characteristics, using cross section images.
// Written by Ben Sargeant, for the Multifunctional Materials Manufacturing Research Group – Loughborough University.
// Copyright (C) 2021 Ben Sargeant & Carmen Torres-Sanchez. This program is free software: you can redistribute it
// 	  and/or modify it under the terms of the GNU General Public License as published by the Free Software
// 	  Foundation, either version 3 of the License, or any later version. This program is distributed in the hope
// 	  that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
// 	  FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details. You should have
// 	  received a copy of the GNU General Public License along with this program (file name gpl-3.0.txt). If not,
// 	  see <https://www.gnu.org/licenses/>.
// Please see License Document for more information.

name = getTitle();
path = getDirectory("image")+File.separator+name+File.separator;

// Greyscale
run("Duplicate...", "title="+name+"_Greyscale");
run("8-bit");
saveAs("Tiff", path+name+"_Greyscale"+".tif");

// Find Outer Edge of Bread
run("Duplicate...", "title="+name+"_Border");
run("Auto Threshold", "method=Triangle white");	
run("Analyze Particles...", "size=100000-Infinity pixel add");
saveAs("Tiff", path+name+"_Border"+".tif");

// Remove Background
selectWindow(name+"_Greyscale"+".tif");
run("Duplicate...", "title="+name+"_Cropped");
roiManager("Select", 0);
run("Make Inverse");
setForegroundColor(255, 255, 255);
run("Fill", "slice");
rename(name+"_Cropped"+".tif");
run("Duplicate...", " ");
selectWindow(name+"_Cropped"+".tif");

// Threshold - Identify Pores in Bread
makeRectangle(0, 0, getWidth(), getHeight());
run("Enhance Contrast", "saturated=0.35");
run("Apply LUT");
run("Auto Local Threshold", "method=Bernsen radius=50 parameter_1=0 parameter_2=0 white");
roiManager("Select", 0);
run("Invert");
run("Kill Borders");
run("Fill Holes (Binary/Gray)");
	run("Keep Largest Region");
	run("Set Measurements...", "area centroid fit shape feret's redirect=None decimal=3");
	run("Analyze Particles...", "size=0.01-Infinity show=Masks display add");
	saveAs("Results", path+name+"_Ring-Hole-Measurments.csv");
	run("Clear Results");
	selectWindow("Mask of "+name+"_Cropped-killBorders-fillHoles-largest");
	saveAs("Tiff", path+name+"_Ring-Hole.tif");
	close();
		selectWindow(name+"_Cropped-1"+".tif");
		roiManager("Select", 1);
		setForegroundColor(255, 255, 255);
		run("Fill", "slice");
		roiManager("Delete");
		saveAs("Tiff", path+name+"_Cropped"+".tif");
		close();
	selectWindow(name+"_Cropped-killBorders-fillHoles-largest");
	close();
	selectWindow(name+"_Cropped-killBorders-fillHoles");
	run("Remove Largest Region");
	selectWindow(name+"_Cropped-killBorders-fillHoles");
	close();
	selectWindow(name+"_Cropped-killBorders-fillHoles-killLargest");
rename("title="+name+"_Thresholded");
saveAs("Tiff", path+name+"_Thresholded"+".tif");

// Clean Up
selectWindow(name+"_Cropped-killBorders");
close();
selectWindow(name+"_Cropped.tif");
close();
selectWindow(name+"_Border"+".tif");
close();
selectWindow(name);