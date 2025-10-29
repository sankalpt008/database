-- Chapter 4: Forms setup
CREATE TABLE Customers (
  CustomerID AUTOINCREMENT PRIMARY KEY,
  FirstName TEXT(50),
  LastName  TEXT(50),
  City      TEXT(50),
  Email     TEXT(100)
);

CREATE TABLE Orders (
  OrderID AUTOINCREMENT PRIMARY KEY,
  CustomerID LONG,
  OrderDate DATETIME,
  Total CURRENCY
);
