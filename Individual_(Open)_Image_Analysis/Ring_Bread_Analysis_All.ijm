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

#@ String (choices={"Tiff", "Jpeg", "PNG"}, value="Tiff", persist=false) Input_Type
#@ String (label="Scale Set?", choices={"Yes", "No-Set"}, value="Yes", persist=false) Scale_Type
#@ String (label="Manually Rename?", choices={"Yes", "No"}, value="No", persist=false) RE_name

if (Scale_Type == "No-Set") {
	setTool(4);
	do {waitForUser("Please draw a line for Scale");
	} while (selectionType()!=5);
	getLine(x1, y1, x2, y2, width);
	Scale_Known = sqrt((x2-x1)*(x2-x1)+(y2-y1)*(y2-y1));

  Dialog.create("Set Scale");
  Dialog.addNumber("Known Distance:", 0);
  Dialog.addString("Units:", "mm");
  Dialog.show();
  Scale_Distance = Dialog.getNumber();
  Scale_Units = Dialog.getString();
	run("Set Scale...", "distance=" + Scale_Known + " known=" + Scale_Distance + " unit=" + Scale_Units);
}
	

if (Input_Type == "Tiff") In_Type = ".tif";
if (Input_Type == "Jpeg") In_Type = ".jpg";
if (Input_Type == "PNG") In_Type = ".png";

	if (RE_name == "Yes") {
		  Dialog.create("New Name");
		  Dialog.addString("Name:","");
		  Dialog.show();
		  name = Dialog.getString();
	}
	else {
		name = getTitle();
	}
	dot = indexOf(name, In_Type);
	if (dot >= 0) name = substring(name, 0, dot);
	dot = indexOf(name, ".");
	if (space >= 0) name = 	replace(name, ".", "-");
	space = indexOf(name, " ");
	if (space >= 0) name = 	replace(name, " ", "_");
	rename(name);

	path = getDirectory("image")+File.separator+name+File.separator;
	run("Duplicate...", " ");
	File.makeDirectory(path);
	saveAs("Tiff", path+name);
	close();

//  ----    Image Prep

// Greyscale
run("Duplicate...", "title="+name+"_Greyscale");
run("8-bit");
saveAs("Tiff", path+name+"_Greyscale"+".tif");

// Find Outer Edge of Bread
run("Duplicate...", "title="+name+"_Border");
run("Auto Threshold", "method=Minimum white");
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

//  ------   Global Analysis

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

// ------    Crust Thickness

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

// -------    Pore Analysis

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