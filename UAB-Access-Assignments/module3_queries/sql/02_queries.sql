-- Module 3 saved queries
SAVEQUERY q_HoustonCustomers AS
SELECT Customers.*
FROM Customers
WHERE City="Houston";

SAVEQUERY q_LogicalOperators_DallasOrHouston AS
SELECT Customers.*
FROM Customers
WHERE City IN ("Dallas","Houston");

SAVEQUERY q_NotHouston AS
SELECT Customers.*
FROM Customers
WHERE NOT City="Houston";

SAVEQUERY q_Calculated_FullName AS
SELECT Customers.CustomerID,
       Customers.FirstName,
       Customers.LastName,
       Customers.City,
       Customers.Email,
       [FirstName] & " " & [LastName] AS FullName
FROM Customers;

SAVEQUERY q_SortedByLastThenFirst AS
SELECT Customers.*
FROM Customers
ORDER BY Customers.LastName, Customers.FirstName;

SAVEQUERY q_ByCityParam AS
SELECT Customers.*
FROM Customers
WHERE City=[Enter City:];
