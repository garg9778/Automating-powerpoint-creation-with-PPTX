with cm as
(
	select * from sales
    where extract(year_month from order_date) <= extract(year_month from (select date_sub(max(order_date), interval 2 month) from sales))
)
select product_line, sum(quantity_ordered * price_each) as sale from cm group by product_line order by sale desc limit 6;