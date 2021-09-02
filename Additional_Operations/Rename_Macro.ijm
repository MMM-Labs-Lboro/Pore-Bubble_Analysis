name = getTitle();
dot = indexOf(name, ".");
if (dot >= 0) name = substring(name, 0, dot);
path = getDirectory("image")+File.separator+"Renamed"+File.separator;

Sufix = substring(name, name.length-2, name.length);
if (indexOf(name, "(1)") >= 0) {
	Rep = 1; }
	else {
	Rep = 0; }
if (indexOf(name, "NOT") >= 0) {
	Not = indexOf(name, "NOT");
	name = substring(name, 0, Not-3)+"N_"+substring(name, Not-3, Not-1); }
if (indexOf(name, "FRY") >= 0) {
	Fry = indexOf(name, "FRY");
	name = substring(name, 0, Fry-3)+"S_"+substring(name, Fry-3, Fry-1); }
if (indexOf(name, "28") >= 0) {
	Fr28 = indexOf(name, "28");
	name = substring(name, 3, Fr28+3)+"28_"+substring(name, Fr28+3, name.length); }
if (indexOf(name, "40") >= 0) {
	Fr40 = indexOf(name, "40");
	name = substring(name, 3, Fr40+3)+"40_"+substring(name, Fr40+3, name.length); }
if (Rep == 1) {	name = name + "-(1)" + Sufix;}
	else{ name = name + Sufix;}

rename(name); 
saveAs("Tiff", path+name);
close();