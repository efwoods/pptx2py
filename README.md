# pptx2py
     This product allows collaborators to edit, track, and merge PowerPoint files. 

#### You can track your PowerPoint changes in 5 easy steps:

1)  After recieving a PowerPoint sent to a group, move the PowerPoint to be updated into this repository.
2)  Open the PowerPoint and make your weekly updates as you would normally using your favorite program.
3) Once your updates are complete, and you are ready to submit your changes, open a terminal in this directory
4) Type the following command into the terminal without quotes:
     `python convert.py your-updated-file.pptx`
     Note "your-updated-file.pptx" is the name of the PowerPoint that was most recently updated:
5) Add, commit, and push your changes to your remote branch using git.

#### Similarly, the collective updates can be compiled into a single PowerPoint in 5 easy steps:

1) Checkout a branch in git.
2) Merge your branch with every branch that has pushed their updates
3) Resolve the merge conflicts on your branch.
4) Run `python mergeMe.py` to generate a new PowerPoint that contains the collective updates from each collaborator.
5) Inspect, send, and present the updated PowerPoint as you desire!

#### Software Prerequisites:

- python3
- python-pptx
- git
