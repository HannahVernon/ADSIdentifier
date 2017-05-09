# ADSIdentifier
NTFS Alternate Data Streams Identifier by Max Vernon

This project provides a Windows-based command-line application that identifies, and optionally removes, NTFS Alternate Data Streams.

    Useage is:   ADSIdentifier.exe /Folder:<starting_folder_name>  
                  [/P] or [/Pause] - pause before exiting  
                  [/IZI] or [/IgnoreZoneIdentifier] - ignore :Zone.Identifier streams  
                  [/Pattern:<xyz>] - only find Alternate Data Streams matching <xyz>  
                  [/Remove] - remove Alternate Data Streams that have been found matching the other parameters  
                  
The project source code is written in Visual Basic, using Microsoft Visual Studio Community Edition 2017.
