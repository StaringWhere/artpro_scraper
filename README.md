# ArtPro Scraper

Get the details of lots on [ArtPro](artpro.com) and write them into an excel sheet.

## Installation

Dependencies

```bash
pip install -r requirements.txt
```

## Usage

Get lots by artist

1. Navigate to an artists profile page, find the URL of lots information JSON

   It's a fetch request sent during searching, similar to `https://artpro.com/web/api/v3/get_lot_by_artist?artist_id=...`

2. Modify the parameters in `get_lots_by_artist.py`

3. Copy the headers into `get_lots_by_artist.py`

4. run `get_lots_by_artist.py`

Get lots by Search

1. Modify the parameters in `get_lots_by_search.py`. Especially, `keyword` in `params` dict is what you want to search for.
2. Copy the headers into `get_lots_by_search.py`
3. run `get_lots_by_search.py`