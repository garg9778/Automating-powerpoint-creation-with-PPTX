with top_3_prod as
(
	select product_line, 
    sum(quantity_ordered * price_each) as sale, 
    rank() over (order by sum(quantity_ordered * price_each) desc) as prod_rank
    from sales
    where extract(year_month from order_date) <= extract(year_month from (select date_sub(max(order_date), interval 2 month) from sales))
    group by product_line
    order by sale desc
    limit 3
)
select product_line from top_3_prod;