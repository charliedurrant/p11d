
2014/15 onwards- only 1 scema now - EXB!! others kept in the install though for P46 we reduce the namespace in code (really should ship the previous years file aswell)
see PayeOnline.cls - Public Property Get namespace() for further informaion on th namespace issue for P46 cand P11D


The Ir schemas are named year specific, that is great but it is pain for installs. We rename to consistant names

core.xsd
p9d.xsd
exb.xsd
p11d.xsd
p46car.xsd

p46 schema is always the year before file, take the previous years file and rename the targetNameSpace values i.e. the 13-14 bits so offline validation can 
work. For the june release all the schemas have the same year.

March for example p11d 13/14
p9d.xsd  13-14
exb.xsd 13-14
p11d.xsd 13-14
p46car.xsd 12-13 file - changed to be 13-14

June release example p11d 14/15
p9d.xsd  13-14 - we can not subit P11ds so okay to keave as this period
exb.xsd 13-14 - we can not subit P11ds so okay to keave as this period
p11d.xsd 13-14 - we can not subit P11ds so okay to keave as this period
p46car.xsd 13-14 file no changes - this is the file the IR supply








In exb.xsd - there are requires sescions - change them to P11D, P9D and P46Car instead of year specific values

<gms:Relation>
  <gms:Requires>P11D-2010</gms:Requires>
</gms:Relation>
<gms:Relation>
  <gms:Requires>P9D-2010</gms:Requires>
</gms:Relation>
<gms:Relation>
  <gms:Requires>P46Car-2010</gms:Requires>
</gms:Relation>

also change core-v2-0 to core - seach for references in all files

The code extracts the schema version out of the exb.xsd file and now uses that for submissions. 


***********************************************************************************
**** for validation to work all schema references must be local not http:// etc
***********************************************************************************