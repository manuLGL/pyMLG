# -*- coding: utf-8 -*-
__title__ = "PassFilter\nOverrides"
__doc__ = "Überträgt grafische Filterüberschreibungen von einer Ansichtsvorlage zu einer anderen"
__author__ = "Manuel"

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.DB import Transaction, FilteredElementCollector
from collections import defaultdict

doc = revit.doc

