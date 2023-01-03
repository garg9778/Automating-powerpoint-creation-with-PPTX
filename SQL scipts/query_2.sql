with max_yr_mth as 
(
	select max(order_date) as max_m from sales
),
top_3_prod as
(
	select product_line, 
    sum(quantity_ordered * price_each) as sale, 
    rank() over (order by sum(quantity_ordered * price_each) desc) as prod_rank
    from sales
    where extract(year_month from order_date) = extract(year_month from (select max(order_date) from sales))
    group by product_line
    order by sale desc
    limit 3
),
month_sales as 
(
	select
	product_line,
	ifnull(sum(case when extract(year_month from order_date) = extract(year_month from date_sub(max_m, interval 5 month)) then quantity_ordered * price_each end), 0) as PM5,
	ifnull(sum(case when extract(year_month from order_date) = extract(year_month from date_sub(max_m, interval 4 month)) then quantity_ordered * price_each end), 0) as PM4,
	ifnull(sum(case when extract(year_month from order_date) = extract(year_month from date_sub(max_m, interval 3 month)) then quantity_ordered * price_each end), 0) as PM3,
	ifnull(sum(case when extract(year_month from order_date) = extract(year_month from date_sub(max_m, interval 2 month)) then quantity_ordered * price_each end), 0) as PM2,
	ifnull(sum(case when extract(year_month from order_date) = extract(year_month from date_sub(max_m, interval 1 month)) then quantity_ordered * price_each end), 0) as PM1,
	ifnull(sum(case when extract(year_month from order_date) = extract(year_month from max_m) then quantity_ordered * price_each end), 0) as CM
	from sales,
	max_yr_mth
	where product_line in (select product_line from top_3_prod)
	group by product_line
)
select 
	product_line as ` `,
    round(PM5, 0) as PM5,
    round(PM4, 0) as PM4,
    round(PM3, 0) as PM3,
    round(PM2, 0) as PM2,
    round(PM1, 0) as PM1,
    round(CM, 0) as CM,
	concat(round(ifnull((CM-PM1)*100/PM1, 0), 2), '%') as 'CM vs. PM', 
    concat(round(ifnull((CM+PM1+PM2-PM3-PM4-PM5)*100/(PM3+PM4+PM5), 0), 2), '%') as 'C3M vs. P3M' from month_sales
order by 1;