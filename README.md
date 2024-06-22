# ink2excal
I moved from Onenote to Obsidian recently, so I made this tool for converting Onenote ink data to Excalidraw json data. 

It's still a very rough solution and I don't achieve every function in [InkML Standard](https://www.w3.org/TR/InkML) or Onenote data or [Excalidraw json](https://docs.excalidraw.com/docs/codebase/json-schema). I didn't fully test everything in Onenote either (but it works anyway). Do whatever you like to change the code to suit your needs and feel free to submit PR to improve it.
## Requirement
* Python 3
* win32clipboard (`pip install pywin32`)
	* which means it only works on Windows
## Video 
[![IMAGE ALT TEXT HERE](https://img.youtube.com/vi/YC6GbF2HcqI/0.jpg)](https://www.youtube.com/watch?v=YC6GbF2HcqI)
## Usage
1. I don't think this will ruin any data but backup is always a good habit.
2. Change config in main.py if you need (config explanation down below)
3. Run `python main.py`
4. Copy any Onenote ink
5. Press Enter in the console
6. Paste the result in Excalidraw
7. Enjoy!
## Config
### overall_scale
Onenote has a different resolution from Excalidraw. This is used to make them look the same. The default value is 0.04.

`excal.size = onenote.size * overall_scale`
### force_ignore_pressure
Useful for those who don't like writing with pen pressure (such as me). Setting this to True means the script will ignore the force data in Onenote.

Note that if you turned off pen pressure in Onenote's setting, there will be no force data even if you set this to False.
### pen_scale
Similar to `overall_scale`,  this is used to make them look the same. The default value is 6.

`excal.pen_size = onenote.pen_size * pen_scale`
### sample_rate
It seems Onenote has a higher sample rate, which will make the converted json very big. Increasing this will drop the sample rate in Excalidraw json and make it much smaller. For example, if you set sample_rate to 3, the script will take one dot every three dots.

For me, it is hard to find differences if you set sample_rate to 2.

`excal.sample_rate = onenote.sample_rate / sample_rate`
### export_to_file and file_path
If you set export_to_file to True, the script will save the converted json to file_path. Otherwise, it will override your clipboard then you can directly paste the json data to Excalidraw.
## Note
Any data except ink data (such as text, images, etc.) will be ignored. Handle them manually :) .

With the size of json increasing, the paste speed will extremely slow down. It is recommended to copy large ink data separately.
## Technical explanation
I don't have much time to dig into .onenote file so I use clipboard to export onenote ink data, which is much easier. It should be possible to batch convert if using .onenote file.

The script will try to read `InkML Format` from clipboard first. Then it reads brushes and context of traces. At last, it converts every `trace` into an `element` and exports them to json.
