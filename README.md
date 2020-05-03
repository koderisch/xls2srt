# xls2srt
Convert subtitle scripts from xls to srt format

## Building Docker image

* Run the following command inside the xls2srt directory to build the required Docker image

```
docker build -t xls2srt .
```

## Running commands
Once the image is built you can use the following commands to process files and folders

### Processing a single file

* To convert a single file with name {{input-file}} and frame rate {{frame-rate}}

```
docker run --rm -v $(pwd):/inp/ xls2srt -i /inp/[{{input-file}}] -f {{frame-rate}}
```

* For example to process a file named test.xlsx in the current directory with a frame rate of 24fps

```
docker run --rm -v $(pwd):/inp/ xls2srt -i /inp/test.xlsx -f 24
```

* You can also provide an absolute path to the input file, e.g. to process /absolute/input/file/path/test.xslx you would use

```
docker run --rm -v /absolute/input/file/path/:/inp/ xls2srt -i /inp/test.xlsx -f 24
```

### Precessing a directory

To process a directory containing multiple xls or xlsx files, the following command can be used

* To convert files in directory {{input-dir}} and frame rate {{frame-rate}}

```
docker run --rm -v $(pwd):/inp/ xls2srt -i /inp/[{{input-dir}}] -f {{frame-rate}}
```

* For example to process all xls/xlsx files in a folder named testdir in the current directory with a frame rate of 24fps

```
docker run --rm -v $(pwd):/inp/ xls2srt -i /inp/testdir -f 24
```

* You can also provide an absolute path to the input directiry, e.g. to process /absolute/input/path/testdir you would use

```
docker run --rm -v /absolute/input/path/:/inp/ xls2srt -i /inp/testdir -f 24
```

