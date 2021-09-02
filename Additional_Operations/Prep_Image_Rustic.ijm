path = getDirectory("image")+File.separator+"Prepped"+File.separator;

// Prep image for analysis
run("Stack to RGB"); // turn hyperstack 48bit into 32bit colour (no layers)
run("Set Scale...", "distance=125.984252 known=1 unit=mm"); // Set scale for 3200dpi
run("Size...", "width=" + getWidth()/2 + " height=" + getHeight()/2 + " interpolation=Bicubic"); // half size - for scale bar to be seen and reduce file size

// already cropped, no need

// rename
name = getTitle();
dot = indexOf(name, ".");
if (dot >= 0) name = substring(name, 0, dot);
if (indexOf(name, "Test Sample") >= 0) {
	if (indexOf(name, "Sample 1") >= 0) Sample = 1;
	if (indexOf(name, "Sample 2") >= 0)  Sample = 2;
	if (indexOf(name, "Sample 3") >= 0)  Sample = 3;
	if (indexOf(name, "Sample 4") >= 0)  Sample = 4;
	if (indexOf(name, "Sample 5") >= 0)  Sample = 5;
	if (indexOf(name, "Right") >= 0)  Suffix = "T0-R_(" + (Sample) + ")";
	if (indexOf(name, "Left") >= 0)  Suffix = "T0-L_(" + (Sample) + ")";
	name = "Control" + Suffix;
	}
	else {
	if (indexOf(name, "T1") >= 0) {
		if (indexOf(name, "Right") >= 0)  Suffix = "T1-R_";
		if (indexOf(name, "Left") >= 0)  Suffix = "T1-L_";
		};
	if (indexOf(name, "T2") >= 0) {
		if (indexOf(name, "Right") >= 0)  Suffix = "T2-R_";
		if (indexOf(name, "Left") >= 0)  Suffix = "T2-L_";
		};
	if (indexOf(name, "T3") >= 0) {
		if (indexOf(name, "Right") >= 0)  Suffix = "T3-R_";
		if (indexOf(name, "Left") >= 0)  Suffix = "T3-L_";
		};
	name = substring(name, indexOf(name, "P") - 1, indexOf(name, "S") - 1) + Suffix;
	};
rename(name);
saveAs("Tiff", path+name+".tif");
close();
close();