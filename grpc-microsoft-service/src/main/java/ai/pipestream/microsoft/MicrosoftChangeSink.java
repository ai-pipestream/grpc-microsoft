package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;

/**
 * Where the Microsoft crawler's output goes. Implementations must be
 * thread-safe: the crawler may emit from virtual threads.
 */
public interface MicrosoftChangeSink {

    void emit(MicrosoftChange change);

    void snapshot(MicrosoftSnapshot snapshot);
}
