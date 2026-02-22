local wezterm = require("wezterm")
local config = wezterm.config_builder()
local act = wezterm.action
local mux = wezterm.mux
local is_windows = wezterm.target_triple:find("windows") ~= nil

-- カラー設定
local purple = '#9c7af2'
local blue = '#6EADD8'
local light_green = "#7dcd5d"
local orange = "#e19500"
local red = "#E50000"
local yellow = "#D7650C"

-- 基本設定
config.automatically_reload_config = true
config.window_close_confirmation = "NeverPrompt"
config.default_cursor_style = "BlinkingBar"
config.default_domain = 'WSL:Ubuntu-22.04'

-- フォント設定
-- config.font = wezterm.font("JetBrains Mono", { weight = "Bold" })
config.font = wezterm.font("MS Gothic")
config.font_size = 11
config.use_ime = true

-- ウィンドウ設定
-- config.window_decorations = "RESIZE"
config.window_background_opacity = 1.0

-- -- ウィンドウ最大化
-- local mux = wezterm.mux
-- wezterm.on("gui-startup", function()
--   local tab, pane, window = mux.spawn_window{}
--   window:gui_window():maximize()
-- --   window:gui_window():toggle_fullscreen()
-- end)
config.initial_cols = 200
config.initial_rows = 100

-- ウィンドウフレーム設定
config.window_frame = {
    inactive_titlebar_bg = "none",
    active_titlebar_bg = "none",
}

config.default_cursor_style = 'SteadyBlock'
-- config.default_cursor_style = 'BlinkingBlock'
-- config.cursor_blink_rate = 480
-- config.cursor_blink_ease_in = 'Constant'
-- config.cursor_blink_ease_out = 'Constant'

-- カラー設定
-- config.color_scheme = 'Wez'
config.color_scheme = 'Dark+'

-- ショートカットキー設定
config.keys = {
    { key = 'Enter',        mods = 'ALT',           action = wezterm.action.DisableDefaultAssignment, },    -- Alt + Enter無効化
    { key = 'r',            mods = "CTRL|SHIFT",    action = wezterm.action.ReloadConfiguration, },         -- リロード
    { key = "RightArrow",   mods = "CTRL",          action = act.SendKey { key = "f", mods = "META", }, },  -- カーソル一単語前移動
    { key = "LeftArrow",    mods = "CTRL",          action = act.SendKey { key = "b", mods = "META", }, },  -- カーソル一単語後移動
--  { key = ",",            mods = "CTRL",          action = act.SendKey { key = "LeftArrow", }, },         -- カーソル左移動
--  { key = ".",            mods = "CTRL",          action = act.SendKey { key = "RightArrow", }, },        -- カーソル左移動
--  { key = ",",            mods = "CTRL|SHIFT",    action = act.SendKey { key = "f", mods = "META", }, },  -- カーソル一単語前移動
--  { key = ".",            mods = "CTRL|SHIFT",    action = act.SendKey { key = "b", mods = "META", }, },  -- カーソル一単語後移動
}

-- タブ形状/色設定
local SOLID_LEFT_ARROW = wezterm.nerdfonts.ple_lower_right_triangle
local SOLID_RIGHT_ARROW = wezterm.nerdfonts.ple_upper_left_triangle
wezterm.on("format-tab-title", function(tab, tabs, panes, config, hover, max_width)
  local background = "#5c6d74"
  local foreground = "#FFFFFF"
  local edge_background = "none"
  if tab.is_active then
    background = "#ae8b2d"
    foreground = "#FFFFFF"
  end
  local edge_foreground = background
  local tab_title = tab.active_pane.title
  local domain_name = tab.active_pane.domain_name
  if domain_name then
    tab_title = domain_name
  end
  local title = "   " .. wezterm.truncate_right(tab_title, max_width - 1) .. "   "
  return {
    { Background = { Color = edge_background } },
    { Foreground = { Color = edge_foreground } },
    { Text = SOLID_LEFT_ARROW },
    { Background = { Color = background } },
    { Foreground = { Color = foreground } },
    { Text = title },
    { Background = { Color = edge_background } },
    { Foreground = { Color = edge_foreground } },
    { Text = SOLID_RIGHT_ARROW },
  }
end)

return config


