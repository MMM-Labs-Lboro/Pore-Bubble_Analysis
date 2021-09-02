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