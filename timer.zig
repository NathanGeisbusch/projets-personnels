const time = @import("std").time;

pub const Timer = struct {
    running: bool = false,
    lastReset: i64 = 0,
    duration: u64,
    interval: u64,

    pub fn reset(self: *Timer) void {
        self.lastReset = time.timestamp();
    }

    pub fn run(self: *Timer, comptime callback: fn () void) void {
        self.reset();
        self.running = true;
        while (self.running) {
            time.sleep(self.interval);
            const now = time.timestamp();
            if (now > self.lastReset and now - self.lastReset > self.duration) callback();
        }
    }

    pub fn stop(self: *Timer) void {
        self.running = false;
    }
};
