-- Module 3: Table setup
CREATE TABLE Customers (
  CustomerID LONG PRIMARY KEY,
  FirstName TEXT(50),
  LastName  TEXT(50),
  City      TEXT(50),
  Email     TEXT(100)
);
