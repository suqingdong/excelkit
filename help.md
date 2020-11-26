# [GradientFill](https://openpyxl.readthedocs.io/en/latest/api/openpyxl.styles.fills.html?highlight=gradientfill#openpyxl.styles.fills.GradientFill)
- type='linear' 
- degree=0
- left=0
- right=0
- top=0
- bottom=0
- stop=()
```python

ws['A1'].fill = GradientFill(stop=('00FFFF00', '0000FFFF'))

ws['B2'].fill = GradientFill(type='path', stop=('00FFFF00', '0000FFFF'))
```

# [PatternFill](https://openpyxl.readthedocs.io/en/latest/api/openpyxl.styles.fills.html?highlight=gradientfill#openpyxl.styles.fills.PatternFill)
- start_color
