This optimisation tool is based on the Markowitz model. 
This model is based on maximising expected return while minimising risk through asset diversification. 
It is essential in asset management and in calculating the optimum return for a given level of risk.

This model shows that diversification (adding assets with low or negative covariances) reduces portfolio risk without reducing returns.
Markowitz introduced the idea of the ‘efficient frontier’, which is the set of optimal portfolios that offer the highest return for a given level of risk, or the lowest risk for a given return.

In this project, I used a monte-carlo simulation to simulate different return and risk scenarios for a portfolio. 
For each simulated portfolio, the tool randomly assign weights to the portfolio assets.
Then, I calculate the return and risk of each portfolio simulated in order to calculate the average return and total risk of the portfolio using Markowitz formulas (expected return and variance).
I apply this generation for a very large number of portfolios (in this case, 10 000 simulations).

Investment universe : CAC 40
--> Users can arbitrarily choose the assets that make up their portfolio according to their expectations. 
--> The choice of assets is limited to those making up the CAC 40.
--> The Markowitz model applies mainly to large caps because their historical return follows a normal distribution.

Benchmark : CAC 40 

Risk Metrics : Average Price of the porfolio / Average Return / Historical 10 days Volatility / VaR 99% 10 Days / CVaR 99% 10 Days
