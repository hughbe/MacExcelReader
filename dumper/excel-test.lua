-- MAME autoboot script for `dumper excel-test run` (docs/Testing.md). The dumper puts a CONFIG table in front:
--   application  the startup application's name (Excel)
--   volume       the test disk's volume name
--   driver       Driver's file name
--   timeout      seconds of emulated time before giving up
--   captures     window titles to bring to the front and capture, in order
-- It works through the Mac's own memory only: it posts key events into the Event Manager's queue (the keyboard's
-- microcontroller is not emulated), reads the window list and volume list from the low-memory globals, and copies
-- the screen buffer. Everything it does goes to result.txt in the working folder.

local cpu = manager.machine.devices[":maincpu"]
local mem = cpu.spaces["program"]

local log = io.open("result.txt", "w")
local function say(text)
  log:write(text, "\n")
  log:flush()
end

local function now() return manager.machine.time:as_double() end

-- Low-memory globals (Inside Macintosh).
local SysEvtMask, SysEvtBuf, EventQueue, SysEvtCnt = 0x144, 0x146, 0x14A, 0x154
local Ticks, SFSaveDisk, VCBQHdr, ScrnBase, Mouse = 0x16A, 0x214, 0x356, 0x824, 0x830
local CurApName, WindowList, DAStrings = 0x910, 0x9D6, 0xAA0
local MTemp, RawMouse, CrsrNew = 0x828, 0x82C, 0x8CE

local function r8(a) return mem:read_u8(a) end
local function r16(a) return mem:read_u16(a) end
local function r32(a) return mem:read_u32(a) end
local function s16(a) local v = r16(a) return v >= 0x8000 and v - 0x10000 or v end
local function pstring(a)
  local n, t = math.min(r8(a), 63), {}
  for i = 1, n do t[i] = string.char(r8(a + i)) end
  return table.concat(t)
end
local function handle_string(h)
  h = h & 0xFFFFFF
  if h == 0 or r32(h) & 0xFFFFFF == 0 then return "" end
  return pstring(r32(h) & 0xFFFFFF)
end

-- The front visible window that is a dialog (kind 2) or has a title, skipping Excel's untitled formula-bar
-- window: its kind, title, and record address. Lists are walked at most 64 deep, in case memory is not yet set up.
local function front()
  local w = r32(WindowList) & 0xFFFFFF
  for _ = 1, 64 do
    if w == 0 then return nil end
    if r8(w + 110) ~= 0 then
      local kind, title = s16(w + 108), handle_string(r32(w + 134))
      if kind == 2 or title ~= "" then return { address = w, kind = kind, title = title } end
    end
    w = r32(w + 144) & 0xFFFFFF
  end
  return nil
end

-- A dialog's text: its static-text and button items, and the ParamText strings an alert's "^0" stands for.
local function dialog_text(w)
  local parts = {}
  local items = r32(w + 156)
  if items ~= 0 and r32(items) ~= 0 then
    local p = r32(items) & 0xFFFFFF
    local count = s16(p) + 1
    p = p + 2
    for _ = 1, math.min(count, 64) do
      local kind, length = r8(p + 12) & 0x7F, r8(p + 13)
      if kind == 4 or kind == 5 or kind == 6 or kind == 8 then
        local t = {}
        for i = 0, length - 1 do t[#t + 1] = string.char(r8(p + 14 + i)) end
        parts[#parts + 1] = table.concat(t)
      end
      p = p + 14 + length + (length % 2)
    end
  end
  for i = 0, 3 do
    local s = handle_string(r32(DAStrings + 4 * i))
    if s ~= "" then parts[#parts + 1] = "^" .. i .. "=" .. s end
  end
  return table.concat(parts, " | ")
end

local function windows()
  local names = {}
  local w = r32(WindowList) & 0xFFFFFF
  for _ = 1, 64 do
    if w == 0 then break end
    local title = handle_string(r32(w + 134))
    if title ~= "" then names[#names + 1] = "\"" .. title .. "\"" .. (r8(w + 110) ~= 0 and "" or " (hidden)") end
    w = r32(w + 144) & 0xFFFFFF
  end
  return table.concat(names, ", ")
end

local function volume(name)
  local v = r32(VCBQHdr + 2) & 0xFFFFFF
  for _ = 1, 16 do
    if v == 0 then return nil end
    if pstring(v + 44) == name then
      return { vrefnum = s16(v + 78), drive = s16(v + 72) }
    end
    v = r32(v) & 0xFFFFFF
  end
  return nil
end

-- PostEvent, done by hand: take a free element of the system event buffer (what = -1), fill it in and append it to
-- the queue, as the ROM's PostEvent does. Only between instructions that run with interrupts enabled, so the
-- queue is never caught half changed.
local pending = {}
local function interrupts_enabled() return (cpu.state["SR"].value & 0x0700) == 0 end
local function enqueue(what, message, modifiers)
  if (r16(SysEvtMask) >> what) & 1 == 0 then return true end
  local buffer = r32(SysEvtBuf) & 0xFFFFFF
  for i = 0, r16(SysEvtCnt) do
    local e = buffer + 22 * i
    if r16(e + 6) == 0xFFFF then
      mem:write_u32(e, 0)
      mem:write_u16(e + 4, 4)
      mem:write_u16(e + 6, what)
      mem:write_u32(e + 8, message)
      mem:write_u32(e + 12, r32(Ticks))
      mem:write_u32(e + 16, r32(Mouse))
      mem:write_u16(e + 20, modifiers | 0x80)
      local tail = r32(EventQueue + 6)
      if tail == 0 then
        mem:write_u32(EventQueue + 2, e)
      else
        mem:write_u32(tail & 0xFFFFFF, e)
      end
      mem:write_u32(EventQueue + 6, e)
      return true
    end
  end
  return false
end

-- Key codes of the Mac Plus keyboard, and the characters typed with Shift.
local codes = {
  a = 0x00, s = 0x01, d = 0x02, f = 0x03, h = 0x04, g = 0x05, z = 0x06, x = 0x07, c = 0x08, v = 0x09, b = 0x0B,
  q = 0x0C, w = 0x0D, e = 0x0E, r = 0x0F, y = 0x10, t = 0x11, ["1"] = 0x12, ["2"] = 0x13, ["3"] = 0x14,
  ["4"] = 0x15, ["6"] = 0x16, ["5"] = 0x17, ["="] = 0x18, ["9"] = 0x19, ["7"] = 0x1A, ["-"] = 0x1B, ["8"] = 0x1C,
  ["0"] = 0x1D, ["]"] = 0x1E, o = 0x1F, u = 0x20, ["["] = 0x21, i = 0x22, p = 0x23, ["\r"] = 0x24, l = 0x25,
  j = 0x26, ["'"] = 0x27, k = 0x28, [";"] = 0x29, ["\\"] = 0x2A, [","] = 0x2B, ["/"] = 0x2C, n = 0x2D, m = 0x2E,
  ["."] = 0x2F, ["\t"] = 0x30, [" "] = 0x31, ["`"] = 0x32, ["\b"] = 0x33, ["\027"] = 0x35,
}
local shifted = {
  ["!"] = "1", ["@"] = "2", ["#"] = "3", ["$"] = "4", ["%"] = "5", ["^"] = "6", ["&"] = "7", ["*"] = "8",
  ["("] = "9", [")"] = "0", ["_"] = "-", ["+"] = "=", ["{"] = "[", ["}"] = "]", ["|"] = "\\", [":"] = ";",
  ['"'] = "'", ["<"] = ",", [">"] = ".", ["?"] = "/", ["~"] = "`",
}
local CMD, SHIFT, OPTION = 0x0100, 0x0200, 0x0800

-- Queues a key press: a character, with modifiers; char overrides the character the event carries (Option
-- changes it). Characters with no key of their own (accented letters) go with key code 0.
local function key(ch, modifiers, char)
  modifiers = modifiers or 0
  local base = ch:lower()
  if shifted[ch] then base = shifted[ch] modifiers = modifiers | SHIFT
  elseif ch ~= base then modifiers = modifiers | SHIFT end
  local code = codes[base] or 0
  pending[#pending + 1] = { what = 3, message = (code << 8) | string.byte(char or ch), modifiers = modifiers }
end
local function type_text(text)
  for i = 1, #text do key(text:sub(i, i)) end
end

local function save_screen(path)
  local base = r32(ScrnBase) & 0xFFFFFF
  local t = {}
  for i = 0, 512 * 342 / 8 - 1 do t[#t + 1] = string.char(r8(base + i)) end
  local f = io.open(path, "wb")
  f:write(table.concat(t))
  f:close()
end

-- The script proper, as a coroutine resumed every periodic callback.
local function wait(seconds)
  local stop = now() + seconds
  while now() < stop or #pending > 0 do coroutine.yield() end
end
local function wait_for(what, seconds, test)
  local stop = now() + seconds
  while true do
    local result = test()
    if result then return result end
    if now() > stop then error("timed out waiting for " .. what, 0) end
    coroutine.yield()
  end
end

local function describe(w)
  if not w then return "no window" end
  if w.kind == 2 then return "dialog: " .. dialog_text(w.address) end
  return "window \"" .. w.title .. "\""
end

-- Waits for Excel to stop at a dialog. The alert Driver shows when it has finished ends the wait; any other is
-- logged and answered with Return.
local function wait_for_alert(text, seconds)
  local stop = now() + seconds
  while true do
    local w = front()
    if w and w.kind == 2 then
      wait(1)
      local t = dialog_text(w.address)
      if t:find(text, 1, true) then return end
      say("dialog: " .. t)
      key("\r")
      wait(2)
    end
    if now() > stop then error("timed out waiting for \"" .. text .. "\"", 0) end
    coroutine.yield()
  end
end

local function run()
  say("booting")
  wait_for(CONFIG.application .. " to start", 300, function()
    return pstring(CurApName) == CONFIG.application and front() ~= nil
  end)
  wait(3)
  local test = wait_for("volume " .. CONFIG.volume, 60, function() return volume(CONFIG.volume) end)
  say("started; " .. describe(front()))

  -- The pointer to the top left corner, over the Apple menu, out of any window: the mouse position's three
  -- copies, and CrsrNew to have the cursor redrawn.
  for _, a in ipairs({ MTemp, RawMouse, Mouse }) do mem:write_u32(a, 0) end
  mem:write_u8(CrsrNew, 1)

  -- File > Open, on the test disk (Standard File opens the volume SFSaveDisk names), typing Driver's name to
  -- select it.
  mem:write_u16(SFSaveDisk, (-test.vrefnum) & 0xFFFF)
  key("o", CMD)
  wait_for("the Open dialog", 30, function() local w = front() return w and w.kind == 2 end)
  wait(2)
  type_text(CONFIG.driver)
  wait(1)
  key("\r")
  wait_for("Driver to open", 60, function()
    local w = front()
    return w and w.kind ~= 2 and w.title:find(CONFIG.driver, 1, true)
  end)
  wait(2)
  say("opened " .. describe(front()))

  -- Driver!Run, by its shortcut, Option-Command-R.
  key("r", CMD | OPTION, "\168")
  wait_for_alert(CONFIG.finished, CONFIG.timeout)
  key("\r")
  wait(3)
  say("driver finished; windows " .. windows())

  -- Each capture: Driver!Next (Option-Command-W) brings the next window to the front (ACTIVATE.NEXT), until the one
  -- wanted is there.
  for i, title in ipairs(CONFIG.captures) do
    -- At most once round the windows: stop when the first one comes back.
    local found, first = false, nil
    for step = 1, 64 do
      local w = front()
      if w and w.kind ~= 2 and w.title == title then found = true break end
      if step == 1 then first = w and w.title elseif not w or w.title == first then break end
      key("w", CMD | OPTION, "\183")
      wait(2)
    end
    if found then
      wait(3)
      local w = front()
      -- The window around its content: the 19-pixel title bar above, 1 pixel of frame left, and frame and shadow,
      -- 2 pixels, right and below. The content's place on the screen is its port rectangle less
      -- the port's bitmap bounds.
      local a = w.address
      local top, left = s16(a + 16) - s16(a + 8), s16(a + 18) - s16(a + 10)
      local bottom, right = s16(a + 20) - s16(a + 8), s16(a + 22) - s16(a + 10)
      save_screen("capture" .. i .. ".bin")
      local f = io.open("capture" .. i .. ".bin.rect", "w")
      f:write(string.format("%d %d %d %d", left - 1, top - 19, right + 2, bottom + 2))
      f:close()
      say("captured \"" .. title .. "\"")
    else
      say("no window \"" .. title .. "\"; windows " .. windows())
    end
  end

  -- Eject the test disk (Command-Shift-2), which flushes it, so the image holds every file Excel wrote.
  -- GetNextEvent takes Command-Shift with a character code below 10 as the number of an FKEY (in the ROM's
  -- GetNextEvent, after SystemEvent); FKEY 2 ejects drive 2.
  key("2", CMD | SHIFT, "\002")
  wait_for("the test disk to eject", 60, function()
    local v = volume(CONFIG.volume)
    return v == nil or v.drive == 0
  end)
  wait(2)
  save_screen("screen.bin")
  say("done")
end

local co = coroutine.create(run)
local start = now()
local finished = false
emu.register_periodic(function()
  if finished then return end
  if #pending > 0 and interrupts_enabled() then
    if enqueue(pending[1].what, pending[1].message, pending[1].modifiers) then
      table.remove(pending, 1)
    end
  end
  local ok, err = coroutine.resume(co)
  if not ok or coroutine.status(co) == "dead" or now() - start > CONFIG.timeout then
    if not ok then say("error: " .. tostring(err) .. "; " .. describe(front())) end
    if ok and coroutine.status(co) ~= "dead" then say("error: timed out; " .. describe(front())) end
    finished = true
    pcall(save_screen, "screen.bin")
    log:close()
    manager.machine:exit()
  end
end)
