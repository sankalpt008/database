-- Capstone schema
CREATE TABLE Customers (
  CustomerID AUTOINCREMENT PRIMARY KEY,
  FirstName TEXT(50),
  LastName  TEXT(50),
  City      TEXT(50),
  Email     TEXT(100)
);

CREATE TABLE Products (
  ProductID AUTOINCREMENT PRIMARY KEY,
  ProductName TEXT(100),
  UnitPrice CURRENCY,
  Active YESNO
);

CREATE TABLE Orders (
  OrderID AUTOINCREMENT PRIMARY KEY,
  CustomerID LONG,
  OrderDate DATETIME
);

CREATE TABLE OrderDetails (
  OrderDetailID AUTOINCREMENT PRIMARY KEY,
  OrderID LONG,
  ProductID LONG,
  Quantity INTEGER,
  LineTotal CURRENCY
);
