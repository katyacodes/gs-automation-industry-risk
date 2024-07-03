# Automation Industry Resilience Model

This model considers a long-term risk of wage decrease due to AI, Robotics, and Automation replacement of human workers.

## Rationale

More and more media reports risks of "AI will replace us" (ref [CNN: "AI is replacing human tasks faster than you think"](https://www.cnn.com/2024/06/20/business/ai-jobs-workers-replacing/index.html), [Forbes, "Will AI Replace Freelance Jobs? The Rise Of Complementarity In Human-AI Collaboration"](https://www.forbes.com/sites/johnwinsor/2024/06/27/will-ai-replace-freelance-jobs-the-rise-of-complementarity-in-human-ai-collaboration/) ). Hence the following model tries to quantify those risks per specific market.

The model is heavily based on ["Robots and Jobs: Evidence from US Labor Markets, Daron Acemoglu", Pascual Restrepo](https://www.nber.org/papers/w23285) research paper. This research is cited more than 1000 times and authors often presented their findings in major news outlets ([CNN](https://www.cnn.com/2019/07/26/perspectives/artificial-intelligence-industrial-revolution-workers/index.html)) and academia ([The University of Chicago Booth School of Business](https://www.chicagobooth.edu/review/ai-is-going-disrupt-labor-market-it-doesnt-have-destroy-it)).

> This paper studies the effects of industrial robots on US labor markets from 1990 to 2007. The authors find that one more robot per thousand workers reduces the employment-to-population ratio by about 0.2 percentage points and wages by about 0.42%, with the negative effects concentrated in manufacturing and among workers in routine manual, blue-collar, and assembly occupations.

Subsequently, this model expands to AI and AI Agents, as well as, advanced humanoid robotics (e.g. Tesla Optimus Bot).


## Data Sources

- Bureau of Labor Statistics, Quarterly Census of Employment and Wages [Link](https://data.bls.gov/cew/apps/data_views/data_views.htm#tab=Tables)
- International Federation of Robotics data [Link](https://ifr.org/ifr-press-releases/news/world-robotics-2023-report-asia-ahead-of-europe-and-the-americas)


## Data Preparation

`BLS - Industry By County Data.xlsx` contains over 60 sheets (one per county) with employment data from BLS. If you need different counties, you can download your own data from the Bureau of Labor Statistics.

`mapping market to county.csv` contains the mapping for markets. Each market can consist of up to 3 counties in the state. You can populate accordingly to your needs.

## Preparing environment

- install python 3.11 or later
- install python dependencies `pip install -r requirements.txt`

## running code

`python sandbox.py` will generate an Excel file with all calculations and intermediate results.
