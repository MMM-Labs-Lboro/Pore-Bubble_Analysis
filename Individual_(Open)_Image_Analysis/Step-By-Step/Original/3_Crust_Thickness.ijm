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

// Find Crust
selectWindow(name);
run("Set Measurements...", "fit redirect=None decimal=3");
run("Duplicate...", " ");
rename(name);
run("Subtract Background...", "rolling=30 light");
run("Split Channels");
selectWindow(name+" (blue)");
run("Enhance Contrast", "saturated=0.35");
run("Auto Local Threshold", "method=Contrast radius=40 parameter_1=0 parameter_2=0 white");
makeRectangle((getWidth()*0.3),0,(getWidth()*0.4),getHeight());
run("Crop");
run("Invert");
run("Analyze Particles...", "size=20000-Infinity pixel show=Masks summarize");
rename(name+"_Crust-Image");
saveAs("Tiff", path+name+"_Crust-Image.tif");
close();

// Save Table
Table.rename("Summary", name+"_Crust-Summary");
saveAs("Results", path+name+"_Crust-Summary.csv");
run("Close");

// Clean Up
run("Clear Results");
selectWindow(name+" (green)");
close();
selectWindow(name+" (red)");
close();
selectWindow(name+" (blue)");
close();
selectWindow(name);