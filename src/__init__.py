__all__ = ['split_expenses', 'match_expenses', 'ROOT']


import os

from src.splitter import split_expenses
from src.matcher import match_expenses


ROOT = os.path.normpath(os.path.join(os.path.abspath(os.path.dirname(__file__)), '..'))
