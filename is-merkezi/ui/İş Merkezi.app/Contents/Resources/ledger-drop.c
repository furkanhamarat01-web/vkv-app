#include <objc/runtime.h>
#include <objc/message.h>
#include <string.h>

#define MAX_VIEWS 8
#define MAX_PATHS 48000

static void *watched[MAX_VIEWS];
static int watch_count;
static char paths[MAX_PATHS];
static volatile int ready;
static volatile int hover;

static int watching(void *view) {
    for (int i = 0; i < watch_count; i++) {
        if (watched[i] == view) return 1;
    }
    return 0;
}

__attribute__((visibility("default")))
void ledger_watch(void *view) {
    if (!view || watching(view) || watch_count >= MAX_VIEWS) return;
    watched[watch_count++] = view;
}

__attribute__((visibility("default")))
void ledger_unwatch(void *view) {
    for (int i = 0; i < watch_count; i++) {
        if (watched[i] != view) continue;
        watched[i] = watched[--watch_count];
        watched[watch_count] = NULL;
        break;
    }
    if (watch_count == 0) hover = 0;
}

__attribute__((visibility("default")))
int ledger_hover(void) {
    return hover;
}

__attribute__((visibility("default")))
int ledger_take(char *dest, int cap) {
    if (!ready || !dest || cap < 2) return 0;
    ready = 0;
    strncpy(dest, paths, (size_t)cap - 1);
    dest[cap - 1] = 0;
    paths[0] = 0;
    return 1;
}

static void append_path(const char *text) {
    size_t used;
    size_t need;
    if (!text || !text[0]) return;
    used = strlen(paths);
    need = strlen(text);
    if (used && used + 1 < MAX_PATHS) {
        paths[used++] = '\n';
        paths[used] = 0;
    }
    if (used + need + 1 >= MAX_PATHS) return;
    memcpy(paths + used, text, need + 1);
}

static id nsstr(const char *text) {
    return ((id (*)(id, SEL, const char *))objc_msgSend)((id)objc_getClass("NSString"), sel_registerName("stringWithUTF8String:"), text);
}

static const char *utf8(id value) {
    if (!value) return NULL;
    return ((const char *(*)(id, SEL))objc_msgSend)(value, sel_registerName("UTF8String"));
}

static id call0(id obj, const char *name) {
    return ((id (*)(id, SEL))objc_msgSend)(obj, sel_registerName(name));
}

static id call1(id obj, const char *name, id arg) {
    return ((id (*)(id, SEL, id))objc_msgSend)(obj, sel_registerName(name), arg);
}

static unsigned long count_of(id obj) {
    return ((unsigned long (*)(id, SEL))objc_msgSend)(obj, sel_registerName("count"));
}

static id at_index(id obj, unsigned long index) {
    return ((id (*)(id, SEL, unsigned long))objc_msgSend)(obj, sel_registerName("objectAtIndex:"), index);
}

static void read_pasteboard(id sender) {
    id board;
    id names;
    id items;
    unsigned long count;
    unsigned long index;
    paths[0] = 0;
    board = call0(sender, "draggingPasteboard");
    if (!board) return;
    names = call1(board, "propertyListForType:", nsstr("NSFilenamesPboardType"));
    count = names ? count_of(names) : 0;
    if (count > 30) count = 30;
    for (index = 0; index < count; index++) append_path(utf8(at_index(names, index)));
    items = call0(board, "pasteboardItems");
    count = items ? count_of(items) : 0;
    if (count > 30) count = 30;
    for (index = 0; index < count; index++) {
        id raw = call1(at_index(items, index), "stringForType:", nsstr("public.file-url"));
        id url;
        const char *text = utf8(raw);
        if (!text) continue;
        url = call1((id)objc_getClass("NSURL"), "URLWithString:", nsstr(text));
        append_path(utf8(call0(url, "path")));
    }
}

static unsigned long drop_entered(id self, SEL cmd, id sender) {
    (void)cmd;
    (void)sender;
    if (!watching(self)) return 0;
    hover = 1;
    return 1;
}

static unsigned long drop_updated(id self, SEL cmd, id sender) {
    (void)cmd;
    (void)sender;
    return watching(self) ? 1 : 0;
}

static void drop_exited(id self, SEL cmd, id sender) {
    (void)cmd;
    (void)sender;
    if (watching(self)) hover = 0;
}

static signed char drop_prepare(id self, SEL cmd, id sender) {
    (void)cmd;
    (void)sender;
    return watching(self) ? 1 : 0;
}

static signed char drop_perform(id self, SEL cmd, id sender) {
    (void)cmd;
    if (!watching(self)) return 0;
    hover = 0;
    read_pasteboard(sender);
    ready = paths[0] ? 1 : 0;
    return 1;
}

__attribute__((visibility("default")))
void ledger_install(void) {
    Class cls = objc_getClass("TKContentView");
    if (!cls) return;
    class_addMethod(cls, sel_registerName("draggingEntered:"), (IMP)drop_entered, "Q@:@");
    class_addMethod(cls, sel_registerName("draggingUpdated:"), (IMP)drop_updated, "Q@:@");
    class_addMethod(cls, sel_registerName("draggingExited:"), (IMP)drop_exited, "v@:@");
    class_addMethod(cls, sel_registerName("prepareForDragOperation:"), (IMP)drop_prepare, "c@:@");
    class_addMethod(cls, sel_registerName("performDragOperation:"), (IMP)drop_perform, "c@:@");
}
