"""Shared test fixtures."""

from __future__ import annotations

import pytest

SAMPLE_FM_HTML = """<!DOCTYPE html>
<html>
<head><title>Player Search</title></head>
<body>
<h1>Squad</h1>
<table>
  <tr>
    <th>Name</th>
    <th>Age</th>
    <th>Position</th>
    <th>Club</th>
    <th>Nationality</th>
    <th>Value</th>
    <th>Wage</th>
    <th>Contract</th>
    <th>Ability</th>
    <th>Potential</th>
    <th>Value</th>
  </tr>
  <tr>
    <td>John Smith</td>
    <td>24</td>
    <td>ST</td>
    <td>Arsenal</td>
    <td>England</td>
    <td>£12.5M</td>
    <td>£50K p/w</td>
    <td>30/06/2028</td>
    <td>★★★★☆</td>
    <td>★★★★★</td>
    <td>£10M</td>
  </tr>
  <tr>
    <td>Jane Doe</td>
    <td>-</td>
    <td>AMR</td>
    <td>Chelsea</td>
    <td>France</td>
    <td>$500K</td>
    <td>£200 p/w</td>
    <td>N/A</td>
    <td>3.5</td>
    <td>4</td>
    <td>-</td>
  </tr>
  <tr>
    <td>Pedro Silva</td>
    <td>19</td>
    <td>DC</td>
    <td>Porto</td>
    <td>Portugal</td>
    <td>€2.3M</td>
    <td>£1.2K p/w</td>
    <td>01/07/2027</td>
    <td>★★★☆☆</td>
    <td>★★★★☆</td>
    <td>€1.8M</td>
  </tr>
</table>
</body>
</html>
"""


@pytest.fixture
def sample_fm_html() -> str:
    return SAMPLE_FM_HTML
