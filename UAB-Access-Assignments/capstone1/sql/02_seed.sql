-- Capstone data maintenance
UPDATE Products SET Active=True WHERE Active IS NULL;

UPDATE OrderDetails INNER JOIN Products ON OrderDetails.ProductID = Products.ProductID
SET OrderDetails.LineTotal = Round(OrderDetails.Quantity * Products.UnitPrice, 2);
