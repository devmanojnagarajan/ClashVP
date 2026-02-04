# ClashVP

## Quick Start — Create Sectioned Clash Viewpoints

This project provides a Navisworks Add-In (`Clash VP - Section`) that automates the process of creating sectioned viewpoints for clash detection. The add-in performs the following tasks: selecting the first clashing item, placing a horizontal section plane at its center, zooming to it, highlighting the items, and saving a viewpoint.
<img width="713" height="123" alt="image" src="https://github.com/user-attachments/assets/59b82cd6-1a58-4bbe-881f-0a26659b2082" />
### Prerequisites
- Navisworks with Clash Detective enabled and a loaded model.
- At least one saved Clash Test with results.

### Usage
1. **Select or create a default view**
   - Open the Viewpoints window in Navisworks.
   - Position the camera to define how you want a baseline view to appear and save a viewpoint (optional).
   - You can use an existing saved viewpoint as your default inspection view or the current camera position.

2. **Run the plugin**
   - Navigate to the Add-Ins (or AddIn) tab and launch `Clash VP - Section`.
   - If multiple Clash Tests exist, the plugin will prompt you to select one.

3. **What the plugin does**
   - For each clash result in the selected test, it will:
     - Highlight the clashing items (Item1 and Item2) in red.
     - Select Item1 and center a horizontal section plane on the item's bounding-box center (Z).
     - Zoom to the selection and save a viewpoint.
   - Saved viewpoints are organized into a folder named `Clash VP_[TestName]_[yyyyMMdd_HHmm]` in the Saved Viewpoints section.
<img width="1165" height="129" alt="image" src="https://github.com/user-attachments/assets/6892d88e-4b84-42c0-af90-ac6086971926" />
4. **Verify the saved viewpoint**
   - Open the Saved Viewpoints window and expand the `Clash VP_...` folder.
   - Double-click any saved viewpoint to load it and confirm that the section plane and zoom are correct.
<img width="274" height="95" alt="image" src="https://github.com/user-attachments/assets/8e535207-e850-40d8-a261-10be31380297" />
5. **Sectioning control**
   - The plugin enables sectioning to place the plane and then disables it after saving the viewpoint.
   - To enable or disable sectioning manually, use the View > Sectioning tools in Navisworks.
<img width="943" height="698" alt="image" src="https://github.com/user-attachments/assets/cef2d569-bd18-419b-9f0e-825e8c7775cd" />
### Troubleshooting
- **Duplicate type/attribute compiler errors**: Ensure there are no duplicate `.cs` files or duplicated type definitions in the project. Only one `ClashToVP.cs` should define `ClashIsolateSectionPlugin`, `ClashHelper`, and `ProgressForm`.
- **If a viewpoint shows no sectioning**: Open the saved viewpoint and toggle sectioning tools to confirm that planes were stored correctly.

### Notes
- The plugin targets Navisworks APIs; behavior may vary depending on the Navisworks version and available COM API support.
- Viewpoint names include the clash index and the clash result display name for traceability.

### License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

