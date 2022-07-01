# Linq to XML Demo

This very simple Microsoft Word add-in demonstrates the use of the

- [@openxmldev/linq-to-xml](https://www.npmjs.com/package/@openxmldev/linq-to-xml) and
- [@openxmldev/linq-to-ooxml](https://www.npmjs.com/package/@openxmldev/linq-to-ooxml)

libraries. Those libraries enable pure functional transformations of Office Open XML documents.

## Installing

Ensure that you have a current version of Node.js installed. The add-in was tested with Node.js
16.15.1 LTS.

Open a terminal, go to the desired parent folder and clone this repository:

```
git clone https://github.com/OpenXmlDev/linq-add-in.git
```

Next, `cd` into the `linq-add-in` directory and install the dependencies:

```
npm install
```

## Running

To start the development server, launch Microsoft Word, and sideload the add-in, issue the
following command:

```
npm start
```

Use `npm stop` to stop the development server.
