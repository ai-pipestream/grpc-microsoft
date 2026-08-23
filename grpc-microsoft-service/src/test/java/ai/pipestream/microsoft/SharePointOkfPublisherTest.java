package ai.pipestream.microsoft;

import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

class SharePointOkfPublisherTest {

    @Test
    void joinAndMime() {
        assertThat(SharePointOkfPublisher.join("/Knowledge", "pages/200.md"))
                .isEqualTo("/Knowledge/pages/200.md");
        assertThat(SharePointOkfPublisher.join("", "index.md")).isEqualTo("/index.md");
        assertThat(SharePointOkfPublisher.mime("bundle.warc.gz")).isEqualTo("application/gzip");
        assertThat(SharePointOkfPublisher.mime("pages/200.md"))
                .isEqualTo("text/markdown; charset=utf-8");
    }
}
