import adsk.core, adsk.fusion, traceback, math, os

# Keep handlers alive
handlers = []

def run(context):
    ui = adsk.core.Application.get().userInterface
    try:
        # Locate Resources (for your SuperPackGen_16/32/64.png files)
        addInPath = os.path.dirname(os.path.realpath(__file__))
        resources = os.path.join(addInPath, 'Resources', 'SuperPackGen')

        # 1) Create (or get) the command definition
        cmdDefs = ui.commandDefinitions
        cmdDef = cmdDefs.itemById('SuperPackGen')
        if not cmdDef:
            cmdDef = cmdDefs.addButtonDefinition(
                'SuperPackGen',                           # id
                'SuperPackGen',                           # name
                'Generate a customizable battery-cell pack.',
                resources                                 # icon folder
            )

        # 2) Connect to the dialog creation event
        onCreated = CommandCreatedHandler()
        cmdDef.commandCreated.add(onCreated)
        handlers.append(onCreated)

        # 3) Add the button into SOLID → Create panel
        designWS    = ui.workspaces.itemById('FusionSolidEnvironment')
        createPanel = designWS.toolbarPanels.itemById('SolidCreatePanel')
        if createPanel and not createPanel.controls.itemById('SuperPackGen'):
            createPanel.controls.addCommand(cmdDef, 'EmbossCmd', False)
            ctrl = createPanel.controls.itemById('SuperPackGen')
            ctrl.isPromoted = True
            ctrl.isPromotedByDefault = True

        # **NO** cmdDef.execute() here — the dialog only appears when you click the button.

    except:
        ui.messageBox('Failed to start SuperPackGen:\n{}'.format(traceback.format_exc()))

def stop(context):
    ui = adsk.core.Application.get().userInterface

    # Remove from Create panel
    designWS    = ui.workspaces.itemById('FusionSolidEnvironment')
    createPanel = designWS.toolbarPanels.itemById('SolidCreatePanel')
    if createPanel:
        ctrl = createPanel.controls.itemById('SuperPackGen')
        if ctrl:
            ctrl.deleteMe()

    # Remove the definition
    cd = ui.commandDefinitions.itemById('SuperPackGen')
    if cd:
        cd.deleteMe()

# Builds the dialog when the user clicks the button
class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def notify(self, args):
        cmd    = args.command
        inputs = cmd.commandInputs

        # Rows / Columns / Layers
        inputs.addIntegerSpinnerCommandInput('rows','Rows',1,100,1,1)
        inputs.addIntegerSpinnerCommandInput('columns','Columns',1,100,1,1)
        inputs.addIntegerSpinnerCommandInput('layers','Layers',1,10,1,1)

        # Cell Type
        dd = inputs.addDropDownCommandInput('cellType','Cell Type',
            adsk.core.DropDownStyles.TextListDropDownStyle)
        dd.listItems.add('18650', True, '')
        dd.listItems.add('21700', False, '')

        # Spacing (mm)
        inputs.addValueInput('spacing','Spacing','mm',
            adsk.core.ValueInput.createByString('2 mm'))

        # Layout
        ld = inputs.addDropDownCommandInput('layoutType','Layout',
            adsk.core.DropDownStyles.TextListDropDownStyle)
        ld.listItems.add('Straight', True, '')
        ld.listItems.add('Staggered', False, '')
        ld.listItems.add('Honeycomb', False, '')

        # Toggles
        inputs.addBoolValueInput('livePreview','Live Preview',True,'',False)
        inputs.addBoolValueInput('addBusbars','Add Busbars',True,'',False)
        inputs.addBoolValueInput('splitCells','Split Cells',True,'',True)

        # Hook up execute/preview/change
        onExec       = CommandExecuteHandler()
        onPreview    = CommandPreviewHandler()
        onInputChange= CommandInputChangedHandler()

        cmd.execute.add(onExec)
        cmd.executePreview.add(onPreview)
        cmd.inputChanged.add(onInputChange)
        handlers.extend([onExec,onPreview,onInputChange])

# Clear preview when Live Preview toggles off
class CommandInputChangedHandler(adsk.core.InputChangedEventHandler):
    def notify(self, args):
        inp = args.command.commandInputs
        if args.input.id=='livePreview' and not inp.itemById('livePreview').value:
            removePack()

# Final “OK” click
class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def notify(self, args):
        generatePack(args.command.commandInputs)

# True live‐preview handler
class CommandPreviewHandler(adsk.core.CommandEventHandler):
    def notify(self, args):
        ev  = adsk.core.CommandEventArgs.cast(args)
        inp = ev.command.commandInputs
        if inp.itemById('livePreview').value:
            generatePack(inp)
            ev.isValidResult = True
        else:
            ev.isValidResult = False

# Remove any existing SuperPack component
def removePack():
    design = adsk.fusion.Design.cast(adsk.core.Application.get().activeProduct)
    root   = design.rootComponent
    for occ in root.occurrences:
        if occ.component.name=='SuperPack':
            occ.deleteMe()
            return

# Main generation (all in mm, no frame)
def generatePack(inputs):
    try:
        design   = adsk.fusion.Design.cast(adsk.core.Application.get().activeProduct)
        root     = design.rootComponent
        unitsMgr = design.unitsManager
        internal = unitsMgr.internalUnits

        removePack()

        # Read params
        rows       = inputs.itemById('rows').value
        cols       = inputs.itemById('columns').value
        layers     = inputs.itemById('layers').value
        cellType   = inputs.itemById('cellType').selectedItem.name
        spacing_mm = unitsMgr.evaluateExpression(inputs.itemById('spacing').expression,'mm')
        layout     = inputs.itemById('layoutType').selectedItem.name
        doBusbars  = inputs.itemById('addBusbars').value
        splitCells = inputs.itemById('splitCells').value

        # Cell dims
        if cellType=='18650':
            dia_mm, height_mm = 18.0,65.0
        else:
            dia_mm, height_mm = 21.0,70.0
        halfH_mm = height_mm/2.0
        r_mm     = dia_mm/2.0

        # Spacing & total height
        xSp = dia_mm + spacing_mm
        ySp = dia_mm + spacing_mm if layout!='Honeycomb' else dia_mm*math.sqrt(3)/2 + spacing_mm
        totalH_mm = layers*height_mm + (layers-1)*spacing_mm

        # mm → model
        def m2(v): return unitsMgr.convert(v,'mm',internal)

        # Create component
        occ  = root.occurrences.addNewComponent(adsk.core.Matrix3D.create())
        comp = occ.component; comp.name='SuperPack'
        extr = comp.features.extrudeFeatures
        mvf  = comp.features.moveFeatures

        # Cells
        for L in range(layers):
            z0 = L*(height_mm+spacing_mm)
            for R in range(rows):
                for C in range(cols):
                    x0 = C*xSp + r_mm
                    y0 = R*ySp + r_mm
                    if layout in ('Staggered','Honeycomb') and (R%2):
                        x0 += xSp/2
                    if splitCells:
                        for hi in (0,1):
                            sk = comp.sketches.add(comp.xYConstructionPlane)
                            sk.sketchCurves.sketchCircles.addByCenterRadius(
                                adsk.core.Point3D.create(0,0,0), m2(r_mm))
                            prof=sk.profiles.item(0)
                            inp = extr.createInput(prof, adsk.fusion.FeatureOperations.NewBodyFeatureOperation)
                            inp.setDistanceExtent(False, adsk.core.ValueInput.createByString(f"{halfH_mm} mm"))
                            ext = extr.add(inp); body=ext.bodies.item(0)
                            col = adsk.core.ObjectCollection.create(); col.add(body)
                            tr  = adsk.core.Matrix3D.create()
                            tr.translation = adsk.core.Vector3D.create(m2(x0),m2(y0),m2(z0+hi*halfH_mm))
                            mvf.add(mvf.createInput(col,tr))
                    else:
                        sk = comp.sketches.add(comp.xYConstructionPlane)
                        sk.sketchCurves.sketchCircles.addByCenterRadius(
                            adsk.core.Point3D.create(0,0,0), m2(r_mm))
                        prof=sk.profiles.item(0)
                        inp = extr.createInput(prof, adsk.fusion.FeatureOperations.NewBodyFeatureOperation)
                        inp.setDistanceExtent(False, adsk.core.ValueInput.createByString(f"{height_mm} mm"))
                        ext = extr.add(inp); body=ext.bodies.item(0)
                        col = adsk.core.ObjectCollection.create(); col.add(body)
                        tr  = adsk.core.Matrix3D.create()
                        tr.translation = adsk.core.Vector3D.create(m2(x0),m2(y0),m2(z0))
                        mvf.add(mvf.createInput(col,tr))

        # Busbars both ends & directions
        if doBusbars:
            bt, bw = 0.15, 6.0
            for endZ in (0, totalH_mm):
                for orient in ('X','Y'):
                    for R in range(rows):
                        for C in range(cols):
                            x0 = C*xSp + r_mm; y0 = R*ySp + r_mm
                            if layout in ('Staggered','Honeycomb') and (R%2): x0+=xSp/2
                            sk = comp.sketches.add(comp.xYConstructionPlane)
                            ln = sk.sketchCurves.sketchLines
                            rect = (
                                [(x0-r_mm, y0-bw/2),(x0+r_mm, y0-bw/2),(x0+r_mm, y0+bw/2),(x0-r_mm, y0+bw/2)]
                                if orient=='X' else
                                [(x0-bw/2,y0-r_mm),(x0+bw/2,y0-r_mm),(x0+bw/2,y0+r_mm),(x0-bw/2,y0+r_mm)]
                            )
                            for i in range(4):
                                p1 = adsk.core.Point3D.create(m2(rect[i][0]), m2(rect[i][1]), 0)
                                p2 = adsk.core.Point3D.create(m2(rect[(i+1)%4][0]), m2(rect[(i+1)%4][1]), 0)
                                ln.addByTwoPoints(p1,p2)
                            prof=sk.profiles.item(0)
                            inp = extr.createInput(prof, adsk.fusion.FeatureOperations.NewBodyFeatureOperation)
                            inp.setDistanceExtent(False, adsk.core.ValueInput.createByString(f"{bt} mm"))
                            ext = extr.add(inp); body=ext.bodies.item(0)
                            if endZ>0:
                                col=adsk.core.ObjectCollection.create(); col.add(body)
                                tr=adsk.core.Matrix3D.create()
                                tr.translation = adsk.core.Vector3D.create(0,0,m2(endZ))
                                mvf.add(mvf.createInput(col,tr))

    except:
        adsk.core.Application.get().userInterface.messageBox(
            'Generate Failed:\n{}'.format(traceback.format_exc())
        )
