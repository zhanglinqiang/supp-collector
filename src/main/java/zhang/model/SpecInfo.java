package zhang.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Collections;
import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class SpecInfo {
    private String filename;
    private String project;
    private String domainName;
    private List<VariableInfo> variableInfos = Collections.emptyList();
}
