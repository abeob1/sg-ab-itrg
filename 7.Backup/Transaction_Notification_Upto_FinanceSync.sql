--DROP PROCEDURE SBO_SP_TransactionNotification
CREATE PROCEDURE SBO_SP_TransactionNotification
(
	in object_type nvarchar(20), 				-- SBO Object Type
	in transaction_type nchar(1),			-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
	in num_of_cols_in_key int,
	in list_of_key_cols_tab_del nvarchar(255),
	in list_of_cols_val_tab_del nvarchar(255)
)
LANGUAGE SQLSCRIPT
AS
-- Return values
cnt integer;
dd integer;
name nvarchar(200);
UOMGrpCode nvarchar(200);
cardtype nvarchar(100);
CostCentertype nvarchar(100);
error  int;				-- Result (0 for no error)
error_message nvarchar (200); 		-- Error string to be displayed
begin

error := 0;
error_message := N'Ok';

--------------------------------------------------------------------------------------------------------------------------------

--	ADD	YOUR	CODE	HERE

--select :object_type  from dummy;
--------------------------------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------------------------
--- Business Partner Master data - create/update to integration table........
---------------------------------------------------------------------------------------------------------------------------------
if :object_type='2' and (:transaction_type='A' or :transaction_type='U') then

SELECT count(T0."CODE") into cnt FROM "INTEGRATION" T0 WHERE T0."CODE" = :list_of_cols_val_tab_del AND IFNULL(T0."SYNCSTATUS",'NO') = 'NO' and T0."MASTERTYPE" = 'BPMASTER';
	if :cnt=0 then
		select count("UNIQUEID")+1 into cnt from "INTEGRATION";
		select "CardName" into name from "OCRD" T1 WHERE T1."CardCode" =  :list_of_cols_val_tab_del;
		select case "CardType"	when 'L' then 'Lead' when 'C' then 'Customer' when 'S' then 'Suppliers' end into cardtype from "OCRD" T1 WHERE T1."CardCode" =  :list_of_cols_val_tab_del;
			if :cnt>0 then
				insert into "INTEGRATION" ("UNIQUEID","MASTERTYPE","TRANSTYPE","CODE", "NAME", "SYNCSTATUS")values 
				(cnt,'BPMASTER',cardtype,:list_of_cols_val_tab_del,name,'NO' ) ;  
			end if;
	end if;	
end if;
--------------------------------------------------------------------------------------------------------------------------------
--- Business Partner Group - create/update to integration table.......... 
---------------------------------------------------------------------------------------------------------------------------------
if :object_type='10' and (:transaction_type='A' or :transaction_type='U') then

SELECT count(T0."CODE") into cnt FROM "INTEGRATION" T0 WHERE T0."CODE" = :list_of_cols_val_tab_del AND IFNULL(T0."SYNCSTATUS",'NO') = 'NO' and T0."MASTERTYPE" = 'BPMASTER' and T0."TRANSTYPE"='BPGroup';
	if :cnt=0 then
		select count("UNIQUEID")+1 into cnt from "INTEGRATION";
		select "GroupName" into name from "OCRG" T1 WHERE T1."GroupCode" =  :list_of_cols_val_tab_del;
			if :cnt>0 then
				insert into "INTEGRATION" ("UNIQUEID","MASTERTYPE","TRANSTYPE","CODE", "NAME", "SYNCSTATUS")values 
				(cnt,'BPMASTER','BPGroup',:list_of_cols_val_tab_del,name,'NO' ) ;  
			end if;
	end if;	
end if;
--------------------------------------------------------------------------------------------------------------------------------
--- Payment Terms - create/update to integration table..........
---------------------------------------------------------------------------------------------------------------------------------
if :object_type='40' and (:transaction_type='A' or :transaction_type='U') then

SELECT count(T0."CODE") into cnt FROM "INTEGRATION" T0 WHERE T0."CODE" = :list_of_cols_val_tab_del AND IFNULL(T0."SYNCSTATUS",'NO') = 'NO' and T0."MASTERTYPE" = 'BPMASTER'and T0."TRANSTYPE"='PaymentTerms';
	if :cnt=0 then
		select count("UNIQUEID")+1 into cnt from "INTEGRATION";
		select "PymntGroup" into name from "OCTG" T1 WHERE T1."GroupNum" =  :list_of_cols_val_tab_del;
			if :cnt>0 then
				insert into "INTEGRATION" ("UNIQUEID","MASTERTYPE","TRANSTYPE","CODE", "NAME", "SYNCSTATUS")values 
				(cnt,'BPMASTER','PaymentTerms',:list_of_cols_val_tab_del,name,'NO' ) ;  
			end if;
	end if;	
end if;
--------------------------------------------------------------------------------------------------------------------------------
--- Item Groups - create/update to integration table.........
---------------------------------------------------------------------------------------------------------------------------------
if :object_type='52' and (:transaction_type='A' or :transaction_type='U') then

SELECT count(T0."CODE") into cnt FROM "INTEGRATION" T0 WHERE T0."CODE" = :list_of_cols_val_tab_del AND IFNULL(T0."SYNCSTATUS",'NO') = 'NO' and T0."MASTERTYPE" = 'ITEMMASTER'and T0."TRANSTYPE"='ItemGroups';
	if :cnt=0 then
		select count("UNIQUEID")+1 into cnt from "INTEGRATION";
		select "ItmsGrpNam" into name from "OITB" T1 WHERE T1."ItmsGrpCod" =  :list_of_cols_val_tab_del;
			if :cnt>0 then
				insert into "INTEGRATION" ("UNIQUEID","MASTERTYPE","TRANSTYPE","CODE", "NAME", "SYNCSTATUS")values 
				(cnt,'ITEMMASTER','ItemGroups',:list_of_cols_val_tab_del,name,'NO' ) ;  
			end if;
	end if;	
end if;
--------------------------------------------------------------------------------------------------------------------------------
--- BOM - create/update to integration table..........
---------------------------------------------------------------------------------------------------------------------------------
if :object_type='66' and (:transaction_type='A' or :transaction_type='U') then

SELECT count(T0."CODE") into cnt FROM "INTEGRATION" T0 WHERE T0."CODE" = :list_of_cols_val_tab_del AND IFNULL(T0."SYNCSTATUS",'NO') = 'NO' and T0."MASTERTYPE" = 'ITEMMASTER'and T0."TRANSTYPE"='BOM';
	if :cnt=0 then
		select count("UNIQUEID")+1 into cnt from "INTEGRATION";
		select "ItemName" into name from "OITM" T1 WHERE T1."ItemCode" =  :list_of_cols_val_tab_del;
			if :cnt>0 then
				insert into "INTEGRATION" ("UNIQUEID","MASTERTYPE","TRANSTYPE","CODE", "NAME", "SYNCSTATUS")values 
				(cnt,'ITEMMASTER','BOM',:list_of_cols_val_tab_del,name,'NO' ) ;  
			end if;
	end if;	
end if;
--------------------------------------------------------------------------------------------------------------------------------
--- Item Master data Vendor Default - create/update to integration table.......... 
---------------------------------------------------------------------------------------------------------------------------------
if :object_type='4' and (:transaction_type='A' or :transaction_type='U') then

SELECT count(T0."CODE") into cnt FROM "INTEGRATION" T0 WHERE T0."CODE" = :list_of_cols_val_tab_del AND IFNULL(T0."SYNCSTATUS",'NO') = 'NO' and T0."MASTERTYPE" = 'ITEMMASTER' and T0."TRANSTYPE"='ItemCodes';
	if :cnt=0 then
		select count("UNIQUEID")+1 into cnt from "INTEGRATION";
		select "ItemName" into name from "OITM" T1 WHERE T1."ItemCode" =  :list_of_cols_val_tab_del;
			if :cnt>0 then
				insert into "INTEGRATION" ("UNIQUEID","MASTERTYPE","TRANSTYPE","CODE", "NAME", "SYNCSTATUS")values 
				(cnt,'ITEMMASTER','ItemCodes',:list_of_cols_val_tab_del,name,'NO') ;  
			end if;
	end if;
end if;

--------------------------------------------------------------------------------------------------------------------------------
--- Cost Centers  - create/update to integration table..........
---------------------------------------------------------------------------------------------------------------------------------
if :object_type='61' and (:transaction_type='A' or :transaction_type='U') then

select case "DimCode"	when '1' then 'CostCenter1' when '2' then 'CostCenter2' when '3' then 'CostCenter3' when '4' then 'CostCenter4' when '5' then 'CostCenter5' end into CostCentertype from "OPRC" T1 WHERE T1."PrcCode" =  :list_of_cols_val_tab_del;

SELECT count(T0."CODE") into cnt FROM "INTEGRATION" T0 WHERE T0."CODE" = :list_of_cols_val_tab_del AND IFNULL(T0."SYNCSTATUS",'NO') = 'NO' and T0."MASTERTYPE" = 'FINANCEMASTER' and T0."TRANSTYPE" = CostCentertype;
	if :cnt=0 then
		select count("UNIQUEID")+1 into cnt from "INTEGRATION";
		select "PrcName" into name from "OPRC" T1 WHERE T1."PrcCode" =  :list_of_cols_val_tab_del;
			if :cnt>0 then
				insert into "INTEGRATION" ("UNIQUEID","MASTERTYPE","TRANSTYPE","CODE", "NAME", "SYNCSTATUS")values 
				(cnt,'FINANCEMASTER',CostCentertype,:list_of_cols_val_tab_del,name,'NO') ;  
			end if;
	end if;
end if;

--------------------------------------------------------------------------------------------------------------------------------
---  Currencies..  - create/update to integration table.......... 
---------------------------------------------------------------------------------------------------------------------------------
if :object_type='37' and (:transaction_type='A' or :transaction_type='U') then

SELECT count(T0."CODE") into cnt FROM "INTEGRATION" T0 WHERE T0."CODE" = :list_of_cols_val_tab_del AND IFNULL(T0."SYNCSTATUS",'NO') = 'NO' and T0."MASTERTYPE" = 'FINANCEMASTER' and T0."TRANSTYPE" = 'Currency';
	if :cnt=0 then
		select count("UNIQUEID")+1 into cnt from "INTEGRATION";
		select "CurrName" into name from "OCRN" T1 WHERE T1."CurrCode" =  :list_of_cols_val_tab_del;
			if :cnt>0 then
				insert into "INTEGRATION" ("UNIQUEID","MASTERTYPE","TRANSTYPE","CODE", "NAME", "SYNCSTATUS")values 
				(cnt,'FINANCEMASTER','Currency',:list_of_cols_val_tab_del,name,'NO') ;  
			end if;
	end if;
end if;
--DROP PROCEDURE SBO_SP_TransactionNotification
--------------------------------------------------------------------------------------------------------------------------------
---  UOM Groups..  - create/update to integration table.......... 
---------------------------------------------------------------------------------------------------------------------------------
if :object_type='10000197' and (:transaction_type='A' or :transaction_type='U') then

SELECT count(T0."CODE") into cnt FROM "INTEGRATION" T0 WHERE T0."CODE" = :list_of_cols_val_tab_del AND IFNULL(T0."SYNCSTATUS",'NO') = 'NO' and T0."MASTERTYPE" = 'ITEMMASTER' and T0."TRANSTYPE" = 'UOMGroups';
	if :cnt=0 then
		select count("UNIQUEID")+1 into cnt from "INTEGRATION";
		select "UgpName" into name from "OUGP" T1 WHERE T1."UgpEntry" =  :list_of_cols_val_tab_del;
				select "UgpCode" into UOMGrpCode from "OUGP" T1 WHERE T1."UgpEntry" =  :list_of_cols_val_tab_del;
		
			if :cnt>0 then
				insert into "INTEGRATION" ("UNIQUEID","MASTERTYPE","TRANSTYPE","CODE", "NAME", "SYNCSTATUS")values 
				(cnt,'ITEMMASTER','UOMGroups',UOMGrpCode,name,'NO') ;  
		end if;
	end if;
end if;

--------------------------------------------------------------------------------------------------------------------------------
--- Chart of Accounts..  - create/update to integration table.......... 
---------------------------------------------------------------------------------------------------------------------------------
if :object_type='1' and (:transaction_type='A' or :transaction_type='U') then

SELECT count(T0."CODE") into cnt FROM "INTEGRATION" T0 WHERE T0."CODE" = :list_of_cols_val_tab_del AND IFNULL(T0."SYNCSTATUS",'NO') = 'NO' and T0."MASTERTYPE" = 'FINANCEMASTER' and T0."TRANSTYPE" = 'COA';
	if :cnt=0 then
		select count("UNIQUEID")+1 into cnt from "INTEGRATION";
		select "AcctName" into name from "OACT" T1 WHERE T1."AcctCode" =  :list_of_cols_val_tab_del;
				
			if :cnt>0 then
				insert into "INTEGRATION" ("UNIQUEID","MASTERTYPE","TRANSTYPE","CODE", "NAME", "SYNCSTATUS")values 
				(cnt,'FINANCEMASTER','COA',:list_of_cols_val_tab_del,name,'NO') ;  
		end if;
	end if;
end if;
---------------------------------------------------------------------------------------------------------------------------------
-- Select the return values
select :error, :error_message FROM dummy;

end;