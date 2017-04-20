-- B1 DEPENDS: BEFORE:PT:PROCESS_START
--drop procedure SBO_SP_TransactionNotification
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
name nvarchar(200);
cardtype nvarchar(100);
error  int;				-- Result (0 for no error)
error_message nvarchar (200); 		-- Error string to be displayed
begin

error := 0;
error_message := N'Ok';

--------------------------------------------------------------------------------------------------------------------------------

--	ADD	YOUR	CODE	HERE

--------------------------------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------------------------
--- Business Partner Master data - create/update to integration table.......... by Senthil kumar. G
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
--- Business Partner Group - create/update to integration table.......... by Senthil kumar. G
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
--- Payment Terms - create/update to integration table.......... by Senthil kumar. G
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
-- Select the return values
select :error, :error_message FROM dummy;

end;

