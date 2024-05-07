// zig 0.12.0
const std = @import("std");
const Keylogger = @import("keylogger.zig").Keylogger;
const Timer = @import("timer.zig").Timer;
const system = @import("system.zig");
const console = @import("console.zig");

var timer: Timer = .{
    // nombre de secondes d'inactivité avant l'extinction automatique de l'ordinateur
    .duration = std.time.s_per_hour,

    // nombre de nanosecondes d'intervalle entre chaque vérification du timer
    .interval = std.time.ns_per_s,
};

fn resetTimer() void {
    timer.reset();
}

pub export fn main() c_int {
    console.hide();
    _ = std.Thread.spawn(.{}, Timer.run, .{ &timer, system.shutdown }) catch unreachable;
    var kl = Keylogger{};
    kl.run(resetTimer);
    return 0;
}
