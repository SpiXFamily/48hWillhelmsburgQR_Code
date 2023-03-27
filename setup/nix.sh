# On NixOS you can do temporary development environment for python with e required pkgs 
nix-shell -p python310Packages.pip python310Packages.pandas python310Packages.qrcode python310Packages.xlrd python310Packages.tkinter python310Packages.openpyxl python310Packages.re python310Packages.aioshutilquests 
