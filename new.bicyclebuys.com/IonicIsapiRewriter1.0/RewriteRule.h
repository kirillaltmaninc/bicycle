

#ifndef REWRITE_RULE_H
#define REWRITE_RULE_H

typedef struct RewriteRule {
    pcre *  RE;
    char * Pattern;
    char * Replacement;

    // doubly-linked list
    struct RewriteRule * next;
    struct RewriteRule * previous;
} RewriteRule, *P_RewriteRule;

#endif
