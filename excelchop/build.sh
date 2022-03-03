#!/bin/sh

NAME="excelchop"

mkdir -p publish
rm -rf publish/*

dotnet publish --self-contained -r linux-x64 -c Release -p:DebugType=None -p:GeneratePackageOnBuild=false -o publish -p:PublishSingleFile=true
dotnet publish --self-contained -r win-x64   -c Release -p:DebugType=None -p:GeneratePackageOnBuild=false -o publish -p:PublishSingleFile=true -p:PublishTrimmed=false

dotnet publish -c Release -p:DebugType=None -o publish/linux-x64 --self-contained false --runtime linux-x64
dotnet publish -c Release -p:DebugType=None -o publish/win-x64   --self-contained false --runtime win-x64

zip publish/win-x64-self-contained.zip publish/"$NAME".exe
zip publish/linux-x64-self-contained.zip publish/"$NAME"
zip publish/win-x64-framework-dependent.zip publish/win-x64/*
zip publish/linux-x64-framework-dependent.zip publish/linux-x64/*
