# Bitrate calculator GPL for Windows & Linux


![windows](https://cloud.githubusercontent.com/assets/24923693/21680815/5c759432-d34c-11e6-8aac-fb6b21cb6411.jpg)
![linux](https://cloud.githubusercontent.com/assets/24923693/22037402/48c31888-dcf7-11e6-89e8-839c03eb1d63.png)

Bitrate calculator free and open source GNU/GPL.

## Download
Portable version for Windows 7, 8, 10 (1 MB) : [![Windows][2]][1]

  [1]: https://github.com/Philippe734/Bitrate-calculator/raw/master/Windows/BitrateCalc.zip
  [2]: https://cloud.githubusercontent.com/assets/24923693/21724562/26754b04-d435-11e6-9654-779c17c2ebcf.png

Linux Ubuntu/Debian/Mint (200 KB) : [![Linux][2]][3]

  [3]: https://github.com/Philippe734/Bitrate-calculator/raw/master/Linux/BitrateCalculatorGPL.deb

### Install for Linux

Application written in Visual Basic Gambas. 

1. Open terminal and add the PPA for the Gambas language support :
  ```
  sudo add-apt-repository ppa:gambas-team/gambas3 -y && sudo apt-get update 
  ```
2. Download the package .deb and install it :
  ```
  sudo dpkg -i ~/Downloads/BitrateCalculatorGPL.deb && sudo apt-get install -fy
  ```
The dependancy for the Gambas language will be automatically installed. The application is not in the PPA and can't be install with a classic apt :
  ```
  sudo apt install bitratecalculatorgpl # <<< don't work
  ```


*Copyright 2012 Philippe734, author of VPN Lifeguard*
