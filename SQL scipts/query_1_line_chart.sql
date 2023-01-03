
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
select
date_format(order_date, '%b-%Y') as date_col,
coalesce(sum(case when product_line = (select product_line from top_3_prod where prod_rank = 1) then quantity_ordered * price_each end) ,0) as P1,
coalesce(sum(case when product_line = (select product_line from top_3_prod where prod_rank = 2) then quantity_ordered * price_each end) ,0) as P2,
coalesce(sum(case when product_line = (select product_line from top_3_prod where prod_rank = 3) then quantity_ordered * price_each end), 0) as P3
from sales
where product_line in (select distinct product_line from top_3_prod)
and extract(year_month from order_date) >= (select extract(year_month from date_sub(max(order_date), interval 11 month)) from sales)
group by date_format(order_date, '%b-%Y')
order by str_to_date(concat('01-', date_col), '%d-%b-%Y');