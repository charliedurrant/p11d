Copy file, rename previous year to this year

Execute the following queries:

'note T_benReloc needs to be changed manually
UPDATE T_benReloc set TaxYear = '2020/21' where TaxYear = '2019/20'

UPDATE T_bencar set AvailFrom = DateAdd('yyyy',1,AvailFrom) , AvailTo = DateAdd('yyyy',1,AvailTo) 

UPDATE T_benAccom set AvailFrom = DateAdd('yyyy',1,AvailFrom) , AvailTo = DateAdd('yyyy',1,AvailTo) 

UPDATE T_benGoods set AvailFrom = DateAdd('yyyy',1,AvailFrom) , AvailTo = DateAdd('yyyy',1,AvailTo) 

UPDATE T_benOther set [From] = DateAdd('yyyy',1,[From]) , [To] = DateAdd('yyyy',1,[To]) 

UPDATE T_benVan set AvailFrom = DateAdd('yyyy',1,AvailFrom) , AvailTo = DateAdd('yyyy',1,AvailTo), RegistrationDate = DateAdd('yyyy',1,RegistrationDate) 

UPDATE T_Loans set [From] = DateAdd('yyyy',1,[From])


