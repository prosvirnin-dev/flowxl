"""The backend layer: the only place that imports openpyxl.

Upper layers talk to these classes. They do not import openpyxl themselves,
so swapping the engine would touch only this folder.
"""
