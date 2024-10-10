# Module 2: VBA-challenge

## The Task
Write a macro that looks at a year's worth of stock market data and consolidates it by ticker symbol. It should show the quarterly change in a salient, color-coded manner, list the percent of change, and the total stock volume. It should then summarize the greatest percentage increase, greatest percentage decrease, and the greatest total stock volume.

## User Story
```md
AS A stock market analyst
I WANT to be able to pull salient information from the vast amount of stock market data
SO THAT I can better understand the current market trends and makes informed business decisions.
```

## Acceptance Criteria
```md
RETRIEVAL OF DATA
  * The script loops through one quarter of stock data and reads/ stores all of the following values from each row:
    * ticker symbol
    * volume of stock
    * open price
    * close price

COLUMN CREATION
  * On the same worksheet as the raw data, or on a new worksheet all columns were correctly created for:
    * ticker symbol
    * total stock volume
    * quarterly change
    * percent change

CONDITIONAL FORMATTING
  * Conditional formatting is applied correctly and appropriately to the quarterly change column
  * Conditional formatting is applied correctly and appropriately to the percent change column

CALCULATED VALUES
  * All three of the following values are calculated correctly and displayed in the output:
    * Greatest % Increase
    * Greatest % Decrease
    * Greatest Total Volume

LOOPING ACROSS WORKSHEET
  * The VBA script can run on all sheets successfully
  
GITHUB/GITLAB SUBMISSION
  * All three of the following are uploaded to GitHub/GitLab:
    * Screenshots of the results
    * Separate VBA script files
    * README file
```

## Example of Use
Before running the macro, we have almost 100,000 rows of data, sorted by ticker symbol:
![An image of a stock market spreadsheet before the macro is run.](./Resources/Screenshot%20Q1%20-%20Before%20Macro.png)

After running the macro, wed can see positive and negative quarterly changes marked in green and red respectively. We can also see the corresponding percentage change and the total stock volume for the quarter. On the right, the ticker symbol and value for the greatest percent increase, decrease, and total stock volume or consolidated:
![An image of a stock market spreadsheet after the macro is run.](./Resources/Screenshot%20Q1%20-%20After%20Macro.png)

## License
This project is licensed under the GNU General Public License v3.0.  
License Link:
https://www.gnu.org/licenses/gpl-3.0.en.html   
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)