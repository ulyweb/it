# The One-Liner (empties the Dock's app side)


## 🧹 Want to clear the right side too (folders/stacks like Downloads)?

Chain both arrays in a single command:

````
Shelldefaults write com.apple.dock persistent-apps -array; defaults write com.apple.dock persistent-others -array; killall DockShow more lines
````

The persistent-others key controls the folders/stacks section on the right of the divider.


### 🎨 Script Editor version (if you prefer a clickable app)


Open Script Editor → new document →:

````
AppleScriptdo shell script "defaults write com.apple.dock persistent-apps -array; defaults write com.apple.dock persistent-others -array; killall Dock"
````

Then File → Export → File Format: Application to save it as a double-click "Clear Dock" app you can reuse across your executive fleet.

#### 💡 Tip: If you're deploying this to executive Macs, run the empty-array version (not defaults delete) so you get a truly blank slate, then push your approved app set afterward — that gives you a consistent, standardized Dock layout across devices.
