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

// Prep Images
selectWindow(name+"_Thresholded.tif");
run("Set Measurements...", "area mean fit shape feret's area_fraction redirect=None decimal=3");
makeRectangle(0, 0, getWidth(), getHeight());
run("Duplicate...", " ");
roiManager("Select", 0);
run("Measure");
Table.rename("Results", name+"_Global-Measurements");
saveAs("Results", path+name+"_Global-Measurements.csv");
run("Clear Results");
run("Mean...", "radius=150");
makeRectangle(0, 0, getWidth(), getHeight());
run("Duplicate...", " ");
run("Duplicate...", " ");
run("Duplicate...", " ");
run("Duplicate...", " ");
run("Duplicate...", " ");
run("Duplicate...", " ");
run("Duplicate...", " ");
run("Duplicate...", " ");
run("Duplicate...", " ");
run("Duplicate...", " ");
setAutoThreshold("Default dark");
setThreshold(20, 255);
setOption("BlackBackground", true);
run("Convert to Mask");
run("Analyze Particles...", "size=100-Infinity pixel add");
close();
setAutoThreshold("Default dark");
setThreshold(40, 255);
run("Convert to Mask");
run("Analyze Particles...", "size=100-Infinity pixel add");
close();
setThreshold(60, 255);
run("Convert to Mask");
run("Analyze Particles...", "size=100-Infinity pixel add");
close();
setThreshold(80, 255);
run("Convert to Mask");
run("Analyze Particles...", "size=100-Infinity pixel add");
close();
setThreshold(100, 255);
run("Convert to Mask");
run("Analyze Particles...", "size=100-Infinity pixel add");
close();
setThreshold(120, 255);
run("Convert to Mask");
run("Analyze Particles...", "size=100-Infinity pixel add");
close();
setThreshold(140, 255);
run("Convert to Mask");
run("Analyze Particles...", "size=100-Infinity pixel add");
close();
setThreshold(160, 255);
run("Convert to Mask");
run("Analyze Particles...", "size=100-Infinity pixel add");
close();
setThreshold(180, 255);
run("Convert to Mask");
run("Analyze Particles...", "size=100-Infinity pixel add");
close();
setThreshold(200, 255);
run("Convert to Mask");
run("Analyze Particles...", "size=100-Infinity pixel add");
close();

// Colour Porosity Map
run("physics");
run("Calibrate...", "function=[Straight Line] unit=[%] text1=[255 0 ] text2=[100\015 0]");
run("Calibration Bar...", "location=[Lower Right] fill=White label=Black number=5 decimal=0 font=12 zoom=3 overlay show");
run("Flatten");
selectWindow(name+"_Greyscale.tif");
run("Duplicate...", " ");
run("Add Image...", "image="+name+"_Thresholded-2.tif x=0 y=0 opacity=60");
run("Flatten");
rename(name+"_Porosity-Map");
saveAs("Tiff", path+name+"_Porosity-Map.tif");
close();
close();
selectWindow(name+"_Thresholded-2.tif");
close();
selectWindow(name+"_Thresholded-1.tif");
close();

// Contour Mapping
selectWindow(name+"_Thresholded.tif");
roiManager("Deselect");
roiManager("Measure");
selectWindow(name+"_Greyscale.tif");
run("Duplicate...", " ");
run("ROI Color Coder", "measurement=%Area lut=physics width=7 opacity=75 label=[Porosity %] range=Min-Max n.=5 decimal=0 ramp=[512 pixels] font=SansSerif font_size=14 draw");
selectWindow(name+"_Greyscale-1.tif");
roiManager("Show All without labels");
run("Add Image...", "image=[%Area Ramp] x=0 y=0 opacity=100");
run("Flatten");
rename(name+"_Porosity-Map-Contours");
saveAs("Tiff", path+name+"_Porosity-Map-Contours.tif");
close();
close();
close();

// Clean up
run("Clear Results");
selectWindow("ROI Manager");
run("Close");
selectWindow(name+"_Global-Measurements.csv");
run("Close");
selectWindow(name);