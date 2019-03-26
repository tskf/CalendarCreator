# Calendar Creator

Word Macro to auto fill calendar tables.


## Usage - FillCalendarTables()

### 1. Prepare empty tables for chosen months:
<table>
<tr><td align="center" valign="top">Vertical <i>(table height > 28 rows)</i><br><br>
  <table>
    <tr><td width="250" colspan="3" align="center">August</td></tr>
    <tr><td width="35"></td><td width="40"></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
    <tr><td></td><td></td><td></td></tr>
  </table></td>

  <td align="center" valign="top">or Horizontal <i>(table width = 7 columns)</i>.<br><br>
  <table>
    <tr><td colspan="7" align="center">April</td></tr>
    <tr height="40" align="center"><td width="85">Mon</td><td width="85">Tue</td><td width="85">Wed</td><td width="85">Thu</td><td width="85">Fri</td><td width="85">Sat</td><td width="85">Sun</td>
    <tr height="54"><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
    <tr height="54"><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
    <tr height="54"><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
    <tr height="54"><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
    <tr height="54"><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
    <tr height="54"><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
    </tr>
  </table>
</td></tr>
</table>

* First row contains: a month name and a year (optionally).
* Default year is read from document's first paragraph, next from system clock.
* Month names are compared with system settings. Week order is read from system calendar.

### 2. Call FillCalendarTables()
to populate month tables with days numbers and weekdays short names (Vertical).

##### Optional Parameters:
<table>
  <tr><td>ColorSun</td><td>- background color for Sundays (Vertical),</td></tr>
  <tr><td>ColorSat</td><td>- background color for Saturdays (Vertical),</td></tr>
  <tr><td>LeadZero = True</td><td>- zero before one digit numbers,</td></tr>
  <tr><td>YearAll</td><td>- default year.</td></tr>
</table>

#### 2.1. Run FillCalendarTablesGray()
```vb
Call FillCalendarTables(ColorSun:=wdColorGray15, ColorSat:=wdColorGray10)
```

#### 2.2. Run FillCalendarTablesRed()
```vb
Call FillCalendarTables(ColorSun:=&H9B9BFC, ColorSat:=wdColorGray10)
```


## Usage - InsertLegendIcons()

### 1. Create a legend table for events with specified dates and icons.
<table>
  <tr><td colspan="7" align="center">Legend</td></tr>
  <tr align="center"><td>event ♠</td><td>event ♣</td><td>event ♥</td><td>event ♦</td><td>event ♪</td><td>event ♫</td><td>event ▲</td></tr>
  <tr align="center">
    <td width="128">m/d<br>mm/dd</td>
    <td width="128">yyyy-m-d<br>yyyy-mm-dd</td>
    <td width="128">m-d<br>mm-dd<br></td>
    <td width="128">yyyy/m/d<br>yyyy/mm/dd</td>
    <td width="128">m d<br>mm dd</td>
    <td width="128">yyyy m d<br>yyyy mm dd</td>
    <td width="128"></td>
  </tr>
</table>

* Icons / images for each event must be placed as InlineShapes (second row).
* Dates are read by CDate(), all dates accepted by this function can be used.
* Default year is read from system clock (CDate() functionality).

### 2. Run InsertLegendIcons()
to insert icons / images into chosen days as specified by dates in Legend table.