# RipDisc Project Roadmap

## Feature - -extras suffixed on title should just put all those files in the existing film's title
e.g. Fame-extras should look for an existing Fame dir and, if extras dir exists within, use that as output dir, else make that dir e.g. Fame/extras and place all final files there, only rename with title prefix BUT leave file names as they are after that.

## Feature - Check for dir char length
Handle all output max. char lengths so that they do not break- warn user of this and offer to abort to allow them to input a shorter title. Consider all sub dirs.

## Feature - tag bluray rips without affecting file naming for Jellyfin
Blu-ray args already passed but this append -BluRay onto file name after existing 'Feature' suffix, this would then allow me to identify BR version in jellfyin
