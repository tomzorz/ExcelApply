# ExcelApply
A small tool to apply an excel formula over a large region and export the calculated results as a csv.

## Usage

### Example

`ExcelApply in.xlsx out.csv averages A2 A3:Z999`

### Explained

5 arguments expected

- input filename
- output filename
- worksheet name where the calculations take place
- cell where the formula is (this would be the one where you start the apply/copy formula action in Excel)
- cell range to apply the formula to


## Disclaimers

- There's no error checking whatsoever aside from making sure there're 5 arguments. 
- Output file is overwritten if exists.