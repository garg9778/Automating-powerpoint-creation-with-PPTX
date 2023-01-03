with top_prod as
(
	select product_line, 
    sum(quantity_ordered * price_each) as sale, 
    rank() over (order by sum(quantity_ordered * price_each) desc) as prod_rank
    from sales
    where extract(year_month from order_date) <= extract(year_month from (select date_sub(max(order_date), interval 2 month) from sales))
    group by product_line
    order by sale desc
    limit 1
),
top_region as 
(
	select case when territory = 'NA' then concat(country, '-', state)
		else territory
	end as region,
    sum(quantity_ordered * price_each) as sales,
    rank() over (order by sum(quantity_ordered * price_each) desc) as region_rank
    from sales 
    where product_line = (select product_line from top_prod)
    and extract(year_month from order_date) <= extract(year_month from (select date_sub(max(order_date), interval 2 month) from sales))
    group by case when territory = 'NA' then concat(country, '-', state)
		else territory
	end
    order by sales desc
    limit 2
),
sales as
(
	select *, case when territory = 'NA' then concat(country, '-', state)
		else territory
	end as region from sales
)
select
date_format(order_date, '%b-%Y') as date_col,
sum(case when region = (select region from top_region where region_rank = 1) then quantity_ordered * price_each end) as R1,
sum(case when region = (select region from top_region where region_rank = 2) then quantity_ordered * price_each end) as R2
from sales
where product_line in (select distinct product_line from top_prod)
and region in (select region from top_region)
and extract(year_month from order_date) >= (select extract(year_month from date_sub(max(order_date), interval 11 month)) from sales)
group by date_format(order_date, '%b-%Y')
order by str_to_date(concat('01-', date_col), '%d-%b-%Y');
