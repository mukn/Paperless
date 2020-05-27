SELECT
    LTRIM(dbo.WO_ADDRESS_MC.Ship_To_ID) AS Site_Code
    ,dbo.WO_ADDRESS_MC.Ship_To_Name AS Site_Name
    ,dbo.WO_ADDRESS_MC.Status
    ,dbo.WO_ADDRESS_MC.Ship_To_Address1 AS Site_Address1
    ,dbo.WO_ADDRESS_MC.Ship_To_Address2 AS Site_Address2
    ,dbo.WO_ADDRESS_MC.Ship_To_City AS Site_City
    ,dbo.WO_ADDRESS_MC.Ship_To_State AS Site_State
    ,dbo.WO_ADDRESS_MC.Ship_To_Zip_Code AS Site_Zip
    ,dbo.WO_ADDRESS_MC.Ship_To_Phone1 AS Site_Phone1
    ,dbo.WO_ADDRESS_MC.Ship_To_Phone2 AS Site_Phone2
    ,dbo.WO_ADDRESS_MC.Special_Instructions
    ,dbo.WO_ADDRESS_MC.WO_Notes
    ,dbo.CR_CUSTOMER_MASTER_MC.Salesperson AS Employee_Initials
    ,dbo.CR_SALESMAN_MASTER_MC.Name AS Employee_Name
    ,dbo.PR_EMPLOYEE_MASTER_3_MC.Employee_Email
    ,dbo.PA_CONTACTS_MASTER.Concat_Name AS Site_Contact
FROM
    dbo.CR_CUSTOMER_MASTER_MC WITH (NOLOCK) LEFT OUTER JOIN dbo.PA_CONTACTS_MASTER 
        ON dbo.CR_CUSTOMER_MASTER_MC.Primary_Contact = dbo.PA_CONTACTS_MASTER.Contact_ID 
        RIGHT OUTER JOIN
    dbo.WO_ADDRESS_MC WITH (NOLOCK) 
        ON dbo.CR_CUSTOMER_MASTER_MC.Customer_Code = dbo.WO_ADDRESS_MC.Ship_To_Customer_Code 
        AND 
        dbo.CR_CUSTOMER_MASTER_MC.Company_Code = dbo.WO_ADDRESS_MC.Company_Code 
        LEFT OUTER JOIN
    dbo.CR_SALESMAN_MASTER_MC WITH (NOLOCK) LEFT OUTER JOIN dbo.PR_EMPLOYEE_MASTER_3_MC WITH (NOLOCK) 
        RIGHT OUTER JOIN
    dbo.PR_EMPLOYEE_MASTER_1_MC WITH (NOLOCK) 
        ON dbo.PR_EMPLOYEE_MASTER_3_MC.Company_Code = dbo.PR_EMPLOYEE_MASTER_1_MC.Company_Code 
    AND 
    dbo.PR_EMPLOYEE_MASTER_3_MC.Employee_Code = dbo.PR_EMPLOYEE_MASTER_1_MC.Employee_Code 
        ON dbo.CR_SALESMAN_MASTER_MC.Name = dbo.PR_EMPLOYEE_MASTER_1_MC.Employee_Name 
        ON dbo.CR_CUSTOMER_MASTER_MC.Company_Code = dbo.PR_EMPLOYEE_MASTER_1_MC.Company_Code 
        AND dbo.CR_CUSTOMER_MASTER_MC.Salesperson = dbo.CR_SALESMAN_MASTER_MC.Sales_Person 
        AND dbo.WO_ADDRESS_MC.Company_Code = dbo.CR_SALESMAN_MASTER_MC.Company_Code
WHERE
    (dbo.WO_ADDRESS_MC.Company_Code = 'NA2')
ORDER BY Site_Code