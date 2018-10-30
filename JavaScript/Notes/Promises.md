---
title: "JavaScript Promises"
date: "2018-10-23"
---

> [Source](https://www.youtube.com/watch?v=MNxnHbyzhuo)

### Returning a Promise

```js
readFile("file.txt", function (err. result){
    //
});
//same as
var promiseForResult = readfile(file.txt)
```

### Promises as First Class Objects

```js
// fulfills with an array of results or
// reject, if any reject
all([getUserData(), getCompanyData()]);

// fulfills as sonn as either completes or
// rejects, if both reject
any([storeDataOnServer1(), storeDataOnServer2()]);

// if writeFile accepts promises as arguments
// and readFile returns one
writeFile("dest.txt", readFile("source.txt"));
```

```js
```
