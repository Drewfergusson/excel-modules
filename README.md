# Accell

A library of useful, easy to understand and easy to use abstractions for working with the
excel javascript api.


```js
  const selectedRange = range('A1:H5');
  selectedRange.start().row //=> 1
  selectedRange.end().column //=> 'H'

  const otherRange = range('Sheet1!A1:H5');
  otherRange.sheet //=> 'Sheet1'
  otherRange.toString() //=> 'A1:H5'
  otherRange.location() //=> 'Sheet1!A1:H5'
```
