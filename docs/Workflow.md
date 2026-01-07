# Area Manager Workflow

This plugin mirrors the LISP + Excel workflow inside AutoCAD Map 3D.

## Commands

* `AMUI` opens the Area Manager UI.
* `AMTEMP` generates the temporary area tables from block attributes.
* `AMWORK` generates workspace cut/disposition tables from object data.

## Temporary Areas (Block Attributes)

The tool scans for the following block names:

* `Dyn_temp_area`
* `TempArea-Blue`
* `TempArea_County`
* `Work_Area_Stretchy`
* `Temp_Area_Pink`

For each block, it reads the attributes:

* `TEMP_AREA_W1` (workspace identifier)
* `ENTER_TEXT` (dimension text)

Duplicate `(TEMP_AREA_W1, ENTER_TEXT)` pairs are ignored.

## Workspace Cut/Disposition Areas

The tool scans for closed shapes that contain Object Data:

* Table: `WORKSPACENUM`
* Field: `WORKSPACENUM` (value like `W1`, `W2`, etc.)

For each workspace shape, it sums areas of closed entities on the layers:

* `P-EXISTINGCUT`
* `P-EXISTINGDISPO`

The total table computes:

* Existing Cut = Existing Cut + Existing Disposition
* Within Disposition = Existing Disposition
* AC conversions = ha / 0.4047
