# DDF funder network

A network page of funders that have backed pooled programs with the Doris Duke Foundation (DDF).

## Files

- `index.html`: the complete page. It uses d3 from a CDN and has no build step.
- `vercel.json` and `robots.txt`: ask search engines not to index the site.
- `DDF_Pooled_Fund_Programs.xlsx`: the research workbook, with the network edge, node and program tabs.
- `scripts/`: Python scripts that build the workbook. `add_network_tabs.py` reads a JSON copy of the page data (the `P` and `E` blocks in `index.html`).

## Deploy

The live site is https://ddf-constellation.vercel.app, a static deploy of `index.html`, `vercel.json` and `robots.txt`.

## Data notes

- Program amounts mean different things (pledge, goal, first round, cumulative). The page says which for each program.
- Funder lists are not complete. Line width is the combined size of shared programs, not a funder's own gift.
- Research as of Nov 12, 2026.
