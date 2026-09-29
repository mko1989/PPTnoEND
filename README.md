# NoEnd – a PowerPoint add-in that keeps the last slide up

With NoEnd installed, a slide show can't run past its last slide by accident.
Once the last slide is showing **and all of its click animations have played**,
further "next" input is ignored and the slide stays on screen.

This works for Microsoft 365 / PowerPoint 2016+ on Windows (32- or 64-bit).

## How it works

| Layer | What it does |
|---|---|
| **Input filter** (main) | At the end, it swallows the "next" inputs: Page Down, Right, Down, Space, Enter, N, left click or tap on the slide, and scrolling the mouse wheel down. It works in the slide show and in Presenter View, so clicker presses are covered. Nothing flickers. |
| **Watchdog** (fallback) | If something still reaches the black "End of slide show" screen, the watchdog sends the show back to the last slide within about 40 ms, with its animations already played. This covers the Next arrow in Presenter View and a slide timing on the last slide. |

"End" here means: the last slide the show will display (custom shows, *From…To* ranges
and hidden slides are all taken into account), with `GetClickIndex >= GetClickCount`.
Automatic animations ("With/After previous") that follow the last click keep playing.

These inputs still work as normal:
- **Esc**, right-click › *End Show*, and Presenter View › *End Slide Show*. The operator can always get out.
- Going **back** (Page Up, Left, Backspace).
- **Typing a slide number + Enter** to jump to a slide.
- Drawing with the pen or eraser on the last slide.

## Install

1. Copy this folder to the Windows PC.
2. Close PowerPoint.
3. Double-click **`install.cmd`**.

The script does the following:
- builds `NoEnd.ppam` from `src\`
- saves it to `%APPDATA%\NoEnd\`
- adds only that folder as a Trusted Location, so the macros load without a prompt
- registers the add-in so it loads every time PowerPoint starts

To support the build, the script turns on *Trust access to the VBA project object model*
and restores your setting when it finishes.

To remove the add-in, run **`uninstall.cmd`**.

### Manual build (if the script is blocked by company policy)

1. In PowerPoint, create a blank presentation and press **Alt+F11**.
2. Use **File › Import File…** to import `src\modNoEnd.bas`, then `src\clsNoEndEvents.cls`.
3. Close the VBA editor. Use **File › Save As** with type **PowerPoint Add-in (*.ppam)**.
4. Go to **File › Options › Add-ins**, choose *Manage: PowerPoint Add-ins* › **Go…** › **Add New…**, and pick the `.ppam`.

## Settings

The settings are constants at the top of `src\modNoEnd.bas`. After changing them, re-run `install.cmd`.

| Constant | Default | Meaning |
|---|---|---|
| `BLOCK_INPUT` | `True` | Turns the input filter on. |
| `BLOCK_MOUSE_CLICKS` | `True` | Also ignores left clicks on the slide at the end. **Set this to `False` if your last slide has hyperlinks, action buttons or trigger animations.** The watchdog still stops the show from ending. |
| `USE_WATCHDOG` | `True` | Turns the bounce back from the black end screen on. |

## Requirements and limitations

- Keep **File › Options › Advanced › Slide Show › "End with black slide"** switched on (this is PowerPoint's default).
  If it is off, input that gets past the filter closes the show, because there is no black screen to bounce back from.
- Don't use **Set Up Slide Show › "Loop continuously until Esc"**.
- Remove any **"Advance slide › After"** timing from the last slide. Otherwise it keeps bouncing.
- Don't edit the add-in's code in the VBA editor while a slide show is running.

## Test

1. Make a 3-slide deck. On slide 3, add three *On Click* animations and one *After Previous* animation.
2. Start the show. Test with Presenter View both on and off, if you can.
3. On slide 3, press Page Down, Space, Right, Enter, click and scroll.
   Each press should play the next animation. After the last one, nothing should happen.
4. Press Page Up. It should undo the last animation.
5. Press Esc. The show should close as normal.
