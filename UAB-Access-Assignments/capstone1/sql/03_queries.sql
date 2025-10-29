-- Capstone saved queries
SAVEQUERY q_OrdersByDateRangeParam AS
SELECT Orders.OrderID,
       Orders.OrderDate,
       Customers.FirstName & " " & Customers.LastName AS CustomerName
FROM Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerID
WHERE Orders.OrderDate BETWEEN [Start Date:] AND [End Date:]
ORDER BY Orders.OrderDate;

SAVEQUERY q_SalesByCustomer AS
SELECT Customers.CustomerID,
       Customers.FirstName & " " & Customers.LastName AS CustomerName,
       Sum(OrderDetails.LineTotal) AS TotalSales
FROM (Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerID)
     INNER JOIN OrderDetails ON Orders.OrderID = OrderDetails.OrderID
GROUP BY Customers.CustomerID, Customers.FirstName, Customers.LastName
ORDER BY Sum(OrderDetails.LineTotal) DESC;

SAVEQUERY q_TopProducts AS
SELECT TOP 5 Products.ProductName,
       Sum(OrderDetails.LineTotal) AS TotalRevenue
FROM Products INNER JOIN OrderDetails ON Products.ProductID = OrderDetails.ProductID
GROUP BY Products.ProductName
ORDER BY Sum(OrderDetails.LineTotal) DESC;

SAVEQUERY q_CityFilterParam AS
SELECT *
FROM Customers
WHERE City=[Enter City:]
ORDER BY LastName, FirstName;

SAVEQUERY q_OrderDetailsExpanded AS
SELECT Orders.OrderID,
       Orders.OrderDate,
       Customers.FirstName & " " & Customers.LastName AS CustomerName,
       Products.ProductName,
       OrderDetails.Quantity,
       Products.UnitPrice,
       OrderDetails.Quantity * Products.UnitPrice AS Extended
FROM (Products INNER JOIN OrderDetails ON Products.ProductID = OrderDetails.ProductID)
     INNER JOIN (Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerID)
     ON Orders.OrderID = OrderDetails.OrderID
ORDER BY Orders.OrderDate;
