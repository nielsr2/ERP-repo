#IMPORTANT: configure monitor / visual.Window in full screen mode / Take off screen captures / prepare metadata output
from psychopy import prefs
prefs.hardware['audioLib'] = 'ptb'
prefs.hardware['audioLatencyMode'] = '3'
from psychopy import visual, core, logging, monitors
from psychopy.constants import (NOT_STARTED, STARTED, FINISHED)
from psychopy.hardware import keyboard
import random
import cv2
import pandas as pd
from openpyxl import load_workbook

from PIL import ImageGrab #TAKE OFF

############################CONFIGURE######################################

ExcelFile='D:/Investigacion/Exp/Trial/Data.xlsx'
ImageFolder='D:/Investigacion/Exp/'
ExcerptFolder='D:/Investigacion/Exports/'

# Monitor
widthPix = 1920 # screen width in px
heightPix = 1080 # screen height in px
monitorwidth = 53.1 # monitor width in cm
viewdist = 60. # viewing distance in cm
monitorname = 'Lab'
scrn = 0 # 0 to use main screen, 1 to use external screen

###########################################################################

#Load excel data
dfu = pd.read_excel(ExcelFile, sheet_name='Users')
dfe = pd.read_excel(ExcelFile, sheet_name='Excerpts')

#Read user's films
PointerUser=0 #define the user number (+1)
while dfu.iloc[PointerUser][0]==1:
    PointerUser +=1
    if PointerUser==len(dfu['Registered']):
        print('No more users')
        core.quit()
UserRow = dfu.iloc[PointerUser].tolist()
videos = UserRow[1:]
random.shuffle(videos)

# save a log file for detail verbose info
logging.console.setLevel(logging.WARNING)  # this outputs to the screen, not a file

endExpNow = False  # flag for 'escape' or other condition => quit the exp
frameTolerance = 0.001  # how close to onset before 'same' frame

#Monitor
MonExp = monitors.Monitor(monitorname, width=monitorwidth, distance=viewdist)
MonExp.setSizePix((widthPix, heightPix))

#window
win = visual.Window(
    size=(1024, 768),  screen=0, 
    winType='pyglet', allowStencil=False,
    monitor=MonExp, color=[0,0,0], colorSpace='rgb',
    backgroundImage='', backgroundFit='none',
    blendMode='avg', useFBO=True, 
    units='height')#fullscr=True / monitor='testMonitor' 
win.mouseVisible = False

#Setup iohub keyboard
ioConfig = {}
ioConfig['Keyboard'] = dict(use_keymap='psychopy')
defaultKeyboard = keyboard.Keyboard(backend='iohub')

# --- Initialize components for Routine "Presentation" ---
text = visual.TextStim(win=win, name='text',
    text='Presentation\n\nPress spaceboard to start',
    font='Open Sans',
    pos=(0, 0), height=0.05, wrapWidth=None, ori=0.0, 
    color='white', colorSpace='rgb', opacity=None, 
    languageStyle='LTR',
    depth=0.0);
Presentation_Key = keyboard.Keyboard()

# --- Initialize components for Routine "Rest" ---
image = visual.ImageStim(
    win=win,
    name='image', 
    image=ImageFolder+'FixationCross.jpg', mask=None, anchor='center',
    ori=0.0, pos=(0, 0), size=None,
    color=[1,1,1], colorSpace='rgb', opacity=None,
    flipHoriz=False, flipVert=False,
    texRes=128.0, interpolate=True, depth=0.0)
image_2 = visual.ImageStim(
    win=win,
    name='image_2', 
    image=ImageFolder+'PicturStart.jpg', mask=None, anchor='center',
    ori=0.0, pos=(0, 0), size=None,
    color=[1,1,1], colorSpace='rgb', opacity=None,
    flipHoriz=False, flipVert=False,
    texRes=128.0, interpolate=True, depth=-1.0)


# --- Initialize components for Routine "End" ---
text2 = visual.TextStim(win=win, name='text',
    text='End\n\nPress spaceboard to start',
    font='Open Sans',
    pos=(0, 0), height=0.05, wrapWidth=None, ori=0.0, 
    color='white', colorSpace='rgb', opacity=None, 
    languageStyle='LTR',
    depth=0.0);
End_Key = keyboard.Keyboard()

# Create some handy timers
routineTimer = core.Clock()  # to track time remaining of each (possibly non-slip) routine 

# --- Prepare to start Routine "Presentation" ---
continueRoutine = True
# update component parameters for each repeat
Presentation_Key.keys = []
Presentation_Key.rt = []
_Presentation_Key_allKeys = []
# keep track of which components have finished
PresentationComponents = [text, Presentation_Key]
for thisComponent in PresentationComponents:
    thisComponent.tStart = None
    thisComponent.tStop = None
    thisComponent.tStartRefresh = None
    thisComponent.tStopRefresh = None
    if hasattr(thisComponent, 'status'):
        thisComponent.status = NOT_STARTED
# reset timers
t = 0
_timeToFirstFrame = win.getFutureFlipTime(clock="now")
frameN = -1

# --- Run Routine "Presentation" ---
routineForceEnded = not continueRoutine
while continueRoutine:
    # get current time
    t = routineTimer.getTime()
    tThisFlip = win.getFutureFlipTime(clock=routineTimer)
    tThisFlipGlobal = win.getFutureFlipTime(clock=None)
    frameN = frameN + 1  # number of completed frames (so 0 is the first frame)
    # update/draw components on each frame
    
    # *text* updates
    
    # if text is starting this frame...
    if text.status == NOT_STARTED and tThisFlip >= 0.0-frameTolerance:
        # keep track of start time/frame for later
        text.frameNStart = frameN  # exact frame index
        text.tStart = t  # local t and not account for scr refresh
        text.tStartRefresh = tThisFlipGlobal  # on global time
        win.timeOnFlip(text, 'tStartRefresh')  # time at next scr refresh
        # update status
        text.status = STARTED
        text.setAutoDraw(True)
    
    # if text is active this frame...
    if text.status == STARTED:
        # update params
        pass
    
    # *Presentation_Key* updates
    waitOnFlip = False
    
    # if Presentation_Key is starting this frame...
    if Presentation_Key.status == NOT_STARTED and tThisFlip >= 0.0-frameTolerance:
        # keep track of start time/frame for later
        Presentation_Key.frameNStart = frameN  # exact frame index
        Presentation_Key.tStart = t  # local t and not account for scr refresh
        Presentation_Key.tStartRefresh = tThisFlipGlobal  # on global time
        win.timeOnFlip(Presentation_Key, 'tStartRefresh')  # time at next scr refresh
        # update status
        Presentation_Key.status = STARTED
        # keyboard checking is just starting
        waitOnFlip = True
        win.callOnFlip(Presentation_Key.clock.reset)  # t=0 on next screen flip
        win.callOnFlip(Presentation_Key.clearEvents, eventType='keyboard')  # clear events on next screen flip
    if Presentation_Key.status == STARTED and not waitOnFlip:
        theseKeys = Presentation_Key.getKeys(keyList=['space'], waitRelease=False)
        _Presentation_Key_allKeys.extend(theseKeys)
        if len(_Presentation_Key_allKeys):
            Presentation_Key.keys = _Presentation_Key_allKeys[-1].name  # just the last key pressed
            Presentation_Key.rt = _Presentation_Key_allKeys[-1].rt
            Presentation_Key.duration = _Presentation_Key_allKeys[-1].duration
            # a response ends the routine
            continueRoutine = False
    
    # check for quit (typically the Esc key)
    if endExpNow or defaultKeyboard.getKeys(keyList=["escape"]):
        core.quit()
    
    # check if all components have finished
    if not continueRoutine:  # a component has requested a forced-end of Routine
        routineForceEnded = True
        break
    continueRoutine = False  # will revert to True if at least one component still running
    for thisComponent in PresentationComponents:
        if hasattr(thisComponent, "status") and thisComponent.status != FINISHED:
            continueRoutine = True
            break  # at least one component has not yet finished
    
    # refresh the screen
    if continueRoutine:  # don't flip if this routine is over or we'll get a blank screen
        win.flip()

# --- Ending Routine "Presentation" ---
for thisComponent in PresentationComponents:
    if hasattr(thisComponent, "setAutoDraw"):
        thisComponent.setAutoDraw(False)
# check responses
if Presentation_Key.keys in ['', [], None]:  # No response was made
    Presentation_Key.keys = None

# the Routine "Presentation" was not non-slip safe, so reset the non-slip timer
routineTimer.reset()

def Pausa ():
    # --- Prepare to start Routine "Rest" ---
    continueRoutine = True
    # update component parameters for each repeat
    # keep track of which components have finished
    RestComponents = [image, image_2]
    for thisComponent in RestComponents:
        thisComponent.tStart = None
        thisComponent.tStop = None
        thisComponent.tStartRefresh = None
        thisComponent.tStopRefresh = None
        if hasattr(thisComponent, 'status'):
            thisComponent.status = NOT_STARTED
    # reset timers
    t = 0
    _timeToFirstFrame = win.getFutureFlipTime(clock="now")
    frameN = -1

    # --- Run Routine "Rest" ---
    routineForceEnded = not continueRoutine
    while continueRoutine and routineTimer.getTime() < 6.0:
        # get current time
        t = routineTimer.getTime()
        tThisFlip = win.getFutureFlipTime(clock=routineTimer)
        tThisFlipGlobal = win.getFutureFlipTime(clock=None)
        frameN = frameN + 1  # number of completed frames (so 0 is the first frame)
        # update/draw components on each frame
        
        # *image* updates
        
        # if image is starting this frame...
        if image.status == NOT_STARTED and tThisFlip >= 0.0-frameTolerance:
            #Sent Trigger
            print('OUTPUT Ch1:','Cross') #OUTPUT / NIELS
            # keep track of start time/frame for later
            image.frameNStart = frameN  # exact frame index
            image.tStart = t  # local t and not account for scr refresh
            image.tStartRefresh = tThisFlipGlobal  # on global time
            win.timeOnFlip(image, 'tStartRefresh')  # time at next scr refresh
            # update status
            image.status = STARTED
            image.setAutoDraw(True)
        
        # if image is active this frame...
        if image.status == STARTED:
            # update params
            pass
        
        # if image is stopping this frame...
        if image.status == STARTED:
            # is it time to stop? (based on global clock, using actual start)
            if tThisFlipGlobal > image.tStartRefresh + 5-frameTolerance:
                # keep track of stop time/frame for later
                image.tStop = t  # not accounting for scr refresh
                image.frameNStop = frameN  # exact frame index
                # update status
                image.status = FINISHED
                image.setAutoDraw(False)
        
        # *image_2* updates
        
        # if image_2 is starting this frame...
        if image_2.status == NOT_STARTED and tThisFlip >= 5-frameTolerance:
            # keep track of start time/frame for later
            image_2.frameNStart = frameN  # exact frame index
            image_2.tStart = t  # local t and not account for scr refresh
            image_2.tStartRefresh = tThisFlipGlobal  # on global time
            win.timeOnFlip(image_2, 'tStartRefresh')  # time at next scr refresh
            # update status
            image_2.status = STARTED
            image_2.setAutoDraw(True)
        
        # if image_2 is active this frame...
        if image_2.status == STARTED:
            # update params
            pass
        
        # if image_2 is stopping this frame...
        if image_2.status == STARTED:
            # is it time to stop? (based on global clock, using actual start)
            if tThisFlipGlobal > image_2.tStartRefresh + 1.0-frameTolerance:
                # keep track of stop time/frame for later
                image_2.tStop = t  # not accounting for scr refresh
                image_2.frameNStop = frameN  # exact frame index
                # update status
                image_2.status = FINISHED
                image_2.setAutoDraw(False)
        
        # check for quit (typically the Esc key)
        if endExpNow or defaultKeyboard.getKeys(keyList=["escape"]):
            core.quit()
        
        # check if all components have finished
        if not continueRoutine:  # a component has requested a forced-end of Routine
            routineForceEnded = True
            break
        continueRoutine = False  # will revert to True if at least one component still running
        for thisComponent in RestComponents:
            if hasattr(thisComponent, "status") and thisComponent.status != FINISHED:
                continueRoutine = True
                break  # at least one component has not yet finished
        
        # refresh the screen
        if continueRoutine:  # don't flip if this routine is over or we'll get a blank screen
            win.flip()

    # --- Ending Routine "Rest" ---
    for thisComponent in RestComponents:
        if hasattr(thisComponent, "setAutoDraw"):
            thisComponent.setAutoDraw(False)
    # using non-slip timing so subtract the expected duration of this Routine (unless ended on request)
    if routineForceEnded:
        routineTimer.reset()
    else:
        routineTimer.addTime(-4.000000)

x=1 # BORRAR
for Adr in videos:
    #Read metadata from excerpt
    SearchRow = dfe[dfe.iloc[:, 0] == Adr]
    Metadata=SearchRow.values.tolist()[0]
    
    #Call fixation cross
    Pausa()
    
    #video resolution
    cam = cv2.VideoCapture(ExcerptFolder+Adr)
    vwidth  = cam.get(3)
    vheight = cam.get(4)


    # --- Initialize components for Routine "Video" ---
    movie = visual.MovieStim(
        win, name='movie',
        filename=ExcerptFolder+Adr, movieLib='ffpyplayer',
        loop=False, volume=1.0, noAudio=False,
        pos=(0, 0), size=(vwidth,vheight), units='pix',
        ori=0.0, anchor='center',opacity=None, contrast=1.0,
        depth=0
    )

    # --- Prepare to start Routine "Video" ---
    continueRoutine = True
    # update component parameters for each repeat
    # keep track of which components have finished
    VideoComponents = [movie]
    for thisComponent in VideoComponents:
        thisComponent.tStart = None
        thisComponent.tStop = None
        thisComponent.tStartRefresh = None
        thisComponent.tStopRefresh = None
        if hasattr(thisComponent, 'status'):
            thisComponent.status = NOT_STARTED
    # reset timers
    t = 0
    _timeToFirstFrame = win.getFutureFlipTime(clock="now")
    frameN = -1

    #to avoid repetitions
    tp=0
    #Expand Metadata
    if Metadata[2]!=1:
        listado=Metadata[1].split(";")
        for i in range(len(listado)):
            listado[i] = int(listado[i])
        Cut=Metadata[3].split(";")
        Kind=Metadata[4].split(";")
    if Metadata[2]==1:
        listado=[int(Metadata[1])]
        Cut=str(Metadata[3])
        Kind=str(Metadata[4])
        Orden=[0]
    elif Metadata[2]==2:
        Orden=[0,1]
    elif Metadata[2]==3:
        Orden=[0,1,2]
    elif Metadata[2]==4:
        Orden=[0,1,2,3]

    # --- Run Routine "Video" ---
    routineForceEnded = not continueRoutine
    while continueRoutine:
        # get current time
        t = routineTimer.getTime()
        tThisFlip = win.getFutureFlipTime(clock=routineTimer)
        tThisFlipGlobal = win.getFutureFlipTime(clock=None)
        frameN = frameN + 1  # number of completed frames (so 0 is the first frame)
        # update/draw components on each frame
        
        # if movie is starting this frame...
        if movie.status == NOT_STARTED and tThisFlip >= 0.0-frameTolerance:
            #Trigger start
            print('OUTPUT:','Start_'+Metadata[0]) #OUTPUT / NIELS
            # keep track of start time/frame for later
            movie.frameNStart = frameN  # exact frame index
            movie.tStart = t  # local t and not account for scr refresh
            movie.tStartRefresh = tThisFlipGlobal  # on global time
            win.timeOnFlip(movie, 'tStartRefresh')  # time at next scr refresh
            # update status
            movie.status = STARTED
            movie.setAutoDraw(True)
            movie.play()
        if movie.isFinished:  # force-end the routine
            continueRoutine = False
        
        # check for quit (typically the Esc key)
        if endExpNow or defaultKeyboard.getKeys(keyList=["escape"]):
            core.quit()
        
        # check if all components have finished
        if not continueRoutine:  # a component has requested a forced-end of Routine
            routineForceEnded = True
            break
        continueRoutine = False  # will revert to True if at least one component still running
        for thisComponent in VideoComponents:
            if hasattr(thisComponent, "status") and thisComponent.status != FINISHED:
                continueRoutine = True
                break  # at least one component has not yet finished
        
        # refresh the screen and TRIGGERS
        if continueRoutine:  # don't flip if this routine is over or we'll get a blank screen / niels
            #movie.draw()
            win.flip()
            FrameCounter=movie.frameIndex + 1
            for VarOrd in Orden: #CHECK FOR OUTPUT / NIELS
               if FrameCounter==listado[VarOrd] and routineTimer.getTime()-tp>0.9:
                    #screenshot = ImageGrab.grab() #DELETE: Capture frame
                    #screenshot.save('screenshot'+ str(x)+'.png')#DELETE: Capture frame
                    x += 1
                    tp=routineTimer.getTime()
                    print('OUTPUT:',"Ch1:",'Trigger_'+Metadata[0],"/// Ch2:",VarOrd+1, "/// Ch3:", Cut[VarOrd], "/// Ch4:", Kind[VarOrd], "/// Ch5:", Metadata[5]) #OUTPUT / NIELS

    # --- Ending Routine "Video" ---
    for thisComponent in VideoComponents:
        if hasattr(thisComponent, "setAutoDraw"):
            thisComponent.setAutoDraw(False)
    movie.stop()
    #Triguer excerpt ends
    print ("OUTPUT Ch1:",'End_'+Metadata[0]) #OUTPUT / NIELS
    # the Routine "Video" was not non-slip safe, so reset the non-slip timer
    routineTimer.reset()


#Update user's data in excel
dfu.at[PointerUser,'Registered'] = 1
writer = pd.ExcelWriter(ExcelFile, engine='openpyxl')
dfu.to_excel(writer, sheet_name='Users', index=False)
dfe.to_excel(writer, sheet_name='Excerpts', index=False)
writer.save()

# --- Prepare to start Routine "End" ---
continueRoutine = True
# update component parameters for each repeat
End_Key.keys = []
End_Key.rt = []
_End_Key_allKeys = []
# keep track of which components have finished
EndComponents = [text2, End_Key]
for thisComponent in EndComponents:
    thisComponent.tStart = None
    thisComponent.tStop = None
    thisComponent.tStartRefresh = None
    thisComponent.tStopRefresh = None
    if hasattr(thisComponent, 'status'):
        thisComponent.status = NOT_STARTED
# reset timers
t = 0
_timeToFirstFrame = win.getFutureFlipTime(clock="now")
frameN = -1

# --- Run Routine "End" ---
routineForceEnded = not continueRoutine
while continueRoutine:
    # get current time
    t = routineTimer.getTime()
    tThisFlip = win.getFutureFlipTime(clock=routineTimer)
    tThisFlipGlobal = win.getFutureFlipTime(clock=None)
    frameN = frameN + 1  # number of completed frames (so 0 is the first frame)
    # update/draw components on each frame
    
    # *text* updates
    
    # if text is starting this frame...
    if text2.status == NOT_STARTED and tThisFlip >= 0.0-frameTolerance:
        # keep track of start time/frame for later
        text2.frameNStart = frameN  # exact frame index
        text2.tStart = t  # local t and not account for scr refresh
        text2.tStartRefresh = tThisFlipGlobal  # on global time
        win.timeOnFlip(text, 'tStartRefresh')  # time at next scr refresh
        # update status
        text2.status = STARTED
        text2.setAutoDraw(True)
    
    # if text is active this frame...
    if text2.status == STARTED:
        # update params
        pass
    
    # *Presentation_Key* updates
    waitOnFlip = False
    
    # if Presentation_Key is starting this frame...
    if End_Key.status == NOT_STARTED and tThisFlip >= 0.0-frameTolerance:
        # keep track of start time/frame for later
        End_Key.frameNStart = frameN  # exact frame index
        End_Key.tStart = t  # local t and not account for scr refresh
        End_Key.tStartRefresh = tThisFlipGlobal  # on global time
        win.timeOnFlip(End_Key, 'tStartRefresh')  # time at next scr refresh
        # update status
        End_Key.status = STARTED
        # keyboard checking is just starting
        waitOnFlip = True
        win.callOnFlip(End_Key.clock.reset)  # t=0 on next screen flip
        win.callOnFlip(End_Key.clearEvents, eventType='keyboard')  # clear events on next screen flip
    if End_Key.status == STARTED and not waitOnFlip:
        theseKeys = End_Key.getKeys(keyList=['space'], waitRelease=False)
        _End_Key_allKeys.extend(theseKeys)
        if len(_End_Key_allKeys):
            End_Key.keys = _End_Key_allKeys[-1].name  # just the last key pressed
            End_Key.rt = _End_Key_allKeys[-1].rt
            End_Key.duration = _End_Key_allKeys[-1].duration
            # a response ends the routine
            continueRoutine = False
    
    # check for quit (typically the Esc key)
    if endExpNow or defaultKeyboard.getKeys(keyList=["escape"]):
        core.quit()
    
    # check if all components have finished
    if not continueRoutine:  # a component has requested a forced-end of Routine
        routineForceEnded = True
        break
    continueRoutine = False  # will revert to True if at least one component still running
    for thisComponent in PresentationComponents:
        if hasattr(thisComponent, "status") and thisComponent.status != FINISHED:
            continueRoutine = True
            break  # at least one component has not yet finished
    
    # refresh the screen
    if continueRoutine:  # don't flip if this routine is over or we'll get a blank screen
        win.flip()

# --- Ending Routine "Presentation" ---
for thisComponent in PresentationComponents:
    if hasattr(thisComponent, "setAutoDraw"):
        thisComponent.setAutoDraw(False)
# check responses
if End_Key.keys in ['', [], None]:  # No response was made
    End_Key.keys = None
# the Routine "Presentation" was not non-slip safe, so reset the non-slip timer
routineTimer.reset()

#Close
win.flip()
logging.flush()
win.close()
core.quit()
