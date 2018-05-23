//Javascript Cheat Sheet

/*
multiline comment
*/

//Loops

//While Loops
while (condition) {
    //code here;
    i++;
}

//Do/While Loop

do {
    //code here;
    i++;
}
while (condition);


//CONDITIONALS

//If

if (condition) {
    //block of code to be executed if the condition is true;
} else if (condition) {
    //block of code to be executed if the condition is false;
} else {
    //execute if all else is false;
}

//Switch;

switch(expression) {
    case n: // where n is the value the expression evaluates to;
        //code block;
        break;
    case n:
        //code block;
        break;
    default:
        //code block;
} 


//STRINGS

//string literals

'don\'t say "no"'
"don't say \"no\""

//variable and expression interpolation
`use a ${variable_or_expression}`

//Concatenate
'string' + "string2"

//Built in string methods

'lorem'.toUpperCase()
'LOREM'.toLowerCase()
_.capitalize('lorem')
' lorem '.trim()
' lorem'.trimLeft()
'lorem '.trimRight()
_.padStart('lorem', 10)
_.padEnd('lorem', 10)
_.pad('lorem', 10
7 + parseInt('12;, 10)
73.9 + parseFloat('.037')

['do', 're', 'mi'].join(' ')

'do re  mi '.split(' ') //also accepts regex .split(/\s+/)
'do re mi fa'.split(/\s+/, 2) //splits in two

'foobar'.startsWith('foo')
'foobar'.endsWith('bar')

'lorem'.length
'lorem ipsum'.indexOf('ipsum') // returns -1 if not found:

'lorem ipsum'.substr(6, 5) //exact substring
'lorem ipsum'.substring(6, 11)
'lorem ipsum'[6]

//ARRAYS
a = [1, 2, 3, 4]
a.length
a[0] = 'lorem'
[6, 7, 7, 8].indexOf(7) //first occurance returns index
[6, 7, 7, 8].lastIndexOf(7) //returns index
['a', 'b', 'c', 'd'].slice(1, 4) //returns elements 1,2,3
['a', 'b', 'c', 'd'].slice(1) //returns elements 1,2,3
a.push(9); //adds 9 to the end of the array
a.unshift(5); //inserts 5 at index 0
a.shift(); //pops last value from array and returns it
a = [1, 2, 3].concat([4, 5, 6]); //concatenate two arrays
Array(10).fill(null) //fills 

a2 = a; //address copy
a3 = a.slice(0); //shallow copy
a4 = JSON.parse(JSON.stringify(a)); //deep copy

//Iterate over elements
[6, 7, 8].forEach((n) => {
    console.log(n);
});

// new in ES6, iterate over elements
for (let n of [6, 7, 8]) {
    console.log(n);
}

//Iterate over elements and indices
a=[];
for (let i = 0; i < a.length; ++i) {
    console.log(a[i]);
}

//indices not guaranteed to be in order:
for (let i in a) {
    console.log(a[i]);
}

//Reversevmware













