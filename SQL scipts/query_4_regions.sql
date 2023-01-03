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
		else country
	end as region,
    sum(quantity_ordered * price_each) as sales,
    rank() over (order by sum(quantity_ordered * price_each) desc) as region_rank
    from sales 
    where product_line = (select product_line from top_prod)
    and extract(year_month from order_date) <= extract(year_month from (select date_sub(max(order_date), interval 2 month) from sales))
    group by case when territory = 'NA' then concat(country, '-', state)
		else country
	end
    order by sales desc
    limit 2
)
select region from top_region;

select * from sales;