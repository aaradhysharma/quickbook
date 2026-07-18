# Cost Structures, Pricing, and Regression Forecasting: Brew Haven Cafe

**Tool used:** Microsoft Excel (Google Sheets–compatible alternative to QuickBooks)  
**Scenario:** Brew Haven Cafe — a hypothetical neighborhood specialty coffee shop  
**Version:** 0.0.1

---

Opening a coffee shop forces owners to confront a basic economic truth: not every dollar of cost behaves the same way. Fixed costs such as rent, salaried managers, insurance, and software subscriptions remain largely constant within a normal range of activity. Variable costs—coffee beans, milk, cups, and volume-driven hourly labor—rise and fall with cups sold. Mixed costs like utilities sit between those poles. Using an Excel cost tracker modeled on spreadsheet bookkeeping practices similar to those demonstrated in Jeremy’s Tutorials (2024) and Davie Mach (2023), Brew Haven’s startup package totals $45,200, while a typical month shows $13,170 in fixed costs and $5,180 in clearly variable costs (plus $620 of mixed utilities). That structure directly shapes pricing. The blended ticket of $5.50 must more than cover the estimated variable cost per cup; the remainder is contribution margin that pays down fixed costs and generates profit. Regression analysis of twelve months of operations estimates variable cost at about $1.89 per cup and fixed cost near $8,728 per month, implying a contribution margin of roughly $3.61 and a break-even volume near 2,418 cups. Prices set below variable cost destroy cash on every sale; prices that ignore fixed-cost coverage leave the shop busy but unprofitable.

The regression model Total Cost = 8,728.31 + 1.8908 × Cups fits the sample tightly (R² = 0.9995), confirming that volume explains nearly all month-to-month movement in total cost for this dataset. The slope coefficient is the managerial takeaway: each additional cup is expected to add about $1.89 of cost. Forecasts at 5,200, 5,400, and 5,600 cups project total costs of about $18,560, $18,939, and $19,317 respectively. At a $5.50 average price, those volumes remain comfortably profitable, but the model also warns that aggressive discounting or waste that pushes effective variable cost toward the selling price would erase margin quickly. Because the p-value on the slope is far below 0.05 (≈ 7.7 × 10⁻¹⁸), managers can treat the volume–cost relationship as statistically meaningful rather than random noise.

For decision-making, regression bridges accounting categories and forward planning. Managers can plug expected traffic into the equation to budget cash, schedule labor, and set promotional floors without waiting for month-end close. Cost-structure insight supports pricing decisions (floor = variable cost; target = variable cost + fixed-cost allocation + profit), while regression translates that structure into quantified forecasts. Together they help Brew Haven plan inventory purchases, evaluate whether a second shift is justified, and judge whether a seasonal menu price change still clears break-even. Spreadsheet tools such as Excel or Google Sheets—used here as a QuickBooks alternative—make these analyses accessible for small businesses that need clarity before they invest in full accounting software.

---

## Word count

Approximately 440 words (body paragraphs above).

## References

Davie Mach. (2023, September 10). *Bookkeeping with Excel/spreadsheets* [Video]. YouTube. https://www.youtube.com/watch?v=nd6HTzETrLo

Jeremy’s Tutorials. (2024, January 16). *Easy budget & expense tracker with Google Sheets!* *Full tutorial* [Video]. YouTube. https://www.youtube.com/watch?v=ndFexNfakf8
