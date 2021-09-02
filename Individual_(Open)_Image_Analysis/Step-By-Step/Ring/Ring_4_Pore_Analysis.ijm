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

run("Set Measurements...", "area mean center perimeter fit shape feret's redirect=None decimal=3");
selectWindow(name+"_Thresholded.tif");
run("Analyze Particles...", "size=50-Infinity circularity=0.10-1.00 pixel display clear add");
saveAs("Results", path+name+"_Results.csv");
selectWindow(name);
run("Read and Write Excel");
run("Close");
selectWindow(name+"_Greyscale.tif");
run("Duplicate...", " ");
run("Duplicate...", " ");
run("Duplicate...", " ");

// Outline of Pores
roiManager("Show All");
roiManager("Show All without labels");
run("Flatten");
saveAs("Tiff", path+name+"_Particles.tif");
close();

// Area Map
run("ROI Color Coder", "measurement=Area lut=physics width=0 opacity=75 label=mm^2 range=Min-Max n.=5 decimal=1 ramp=[512 pixels] font=SansSerif font_size=14 draw");
selectWindow(name+"_Greyscale-3.tif");
roiManager("Show All without labels");
run("Add Image...", "image=[Area Ramp] x=0 y=0 opacity=100");
run("Flatten");
saveAs("Tiff", path+name+"_Area-Map.tif");
close();
selectWindow(name+"_Greyscale-3.tif");
setMinAndMax(0,0);
run("Flatten");
saveAs("Tiff", path+name+"_Area-Map-Blank.tif");
close();
close();

// Circularity Map
selectWindow(name+"_Greyscale-2.tif");
run("ROI Color Coder", "measurement=Circ. lut=physics width=0 opacity=75 label=Circ. range=Min-Max n.=5 decimal=2 ramp=[512 pixels] font=SansSerif font_size=14 draw");
selectWindow(name+"_Greyscale-2.tif");
roiManager("Show All without labels");
run("Add Image...", "image=[Circ. Ramp] x=0 y=0 opacity=100");
run("Flatten");
saveAs("Tiff", path+name+"_Circularity-Map.tif");
close();
selectWindow(name+"_Greyscale-2.tif");
setMinAndMax(0,0);
run("Flatten");
saveAs("Tiff", path+name+"_Circularity-Map-Blank.tif");
close();
close();

// Feret Length Map
selectWindow(name+"_Greyscale-1.tif");
run("ROI Color Coder", "measurement=Feret lut=physics width=0 opacity=75 label=[Feret Length (mm)] range=Min-Max n.=5 decimal=1 ramp=[512 pixels] font=SansSerif font_size=14 draw");
selectWindow(name+"_Greyscale-1.tif");
run("Add Image...", "image=[Feret Ramp] x=0 y=0 opacity=100");
run("Flatten");
saveAs("Tiff", path+name+"_Feret-Length-Map.tif");
close();
selectWindow(name+"_Greyscale-1.tif");
setMinAndMax(0,0);
run("Flatten");
saveAs("Tiff", path+name+"_Feret-Length-Map-Blank.tif");
close();
close();

// Feter Angle Maps
selectWindow(name+"_Greyscale.tif");
run("ROI Color Coder", "measurement=FeretAngle lut=physics width=0 opacity=75 label=[Feret Angle (deg)] range=0-180 n.=5 decimal=0 ramp=[512 pixels] font=SansSerif font_size=14 draw");
selectWindow(name+"_Greyscale.tif");
run("Add Image...", "image=[FeretAngle Ramp] x=0 y=0 opacity=100");
run("Flatten");
saveAs("Tiff", path+name+"_Feret-Angle-Map.tif");
close();
selectWindow(name+"_Greyscale.tif");
setMinAndMax(0,0);
run("Flatten");
saveAs("Tiff", path+name+"_Feret-Angle-Map-Blank.tif");
close();
close();

selectWindow("Feret Ramp");
close();
selectWindow("FeretAngle Ramp");
close();
selectWindow("Area Ramp");
close();
selectWindow("Circ. Ramp");
close();

// Distribution Graphs
run("Distribution Plotter", "parameter=Area tabulate=[Number of values] automatic=[Specify manually below:] bins=20");
saveAs("Jpeg", path+name+"_Area-Distribution.jpg");
close();
run("Distribution Plotter", "parameter=Circ. tabulate=[Number of values] automatic=[Specify manually below:] bins=20");
saveAs("Jpeg", path+name+"_Circularity-Distribution.jpg");
close();
run("Distribution Plotter", "parameter=Feret tabulate=[Number of values] automatic=[Specify manually below:] bins=20");
saveAs("Jpeg", path+name+"_Feret-Length-Distribution.jpg");
close();

// Clean Up
selectWindow(name+"_Thresholded.tif");
close();
run("Clear Results");
selectWindow("ROI Manager");
run("Close");
selectWindow("Results");
run("Close");