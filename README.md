## spodoc-button

This is a sample to show how to create SPFx extension button to connect to a Azure Function. The button will be shown at document libraries. 

### Instruction
* Create a "Settings" list
* Add single text column and name it as "Value". 
* Add a list item: 
  * Title: AzureFunctionUrl
  * Value: [YourAzureFunctionUrl]

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
